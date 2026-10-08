import json
import shutil
import subprocess
import unittest
from unittest import mock

from gui import RpaGuiApp


class ErpQueryTest(unittest.TestCase):
    def app(self, attempts):
        app = mock.MagicMock()
        app.is_running = True
        app.config = {'grid_id': '#gridMain'}
        app._float_config = lambda key, default, *args: default
        self.clock = 0.0
        def sleep(seconds):
            self.clock += seconds
            return app.is_running
        app._sleep_interruptible = sleep
        app._wait_for_erp_query_attempt.side_effect = attempts
        return app

    def query(self, app):
        return RpaGuiApp._query_erp_with_retry(
            app, {'search_button': '#gridMain_r'}, '2026-11-15',
            '2026-11-15', 'LJ', '김트랑', 1.0, 0.1,
        )

    def test_completed_empty_then_rows_retries_before_returning(self):
        app = self.app([{'status': 'EMPTY', 'rows': 0}, {'status': 'MATCH', 'rows': 21}])
        self.assertTrue(self.query(app))
        self.assertEqual(app._wait_for_erp_query_attempt.call_count, 2)
        self.assertGreater(self.clock, 0)
        app._save_current_page.assert_not_called()

    def test_three_completed_empty_results_are_no_data(self):
        app = self.app([{'status': 'EMPTY', 'rows': 0}] * 3)
        self.assertFalse(self.query(app))
        self.assertEqual(app._wait_for_erp_query_attempt.call_count, 3)

    def test_timeout_is_failure_without_overlapping_requests(self):
        app = self.app([{'status': 'TIMEOUT', 'rows': 0, 'reason': '응답 시간 초과'}])
        with self.assertRaisesRegex(RuntimeError, '조회 실패.*응답 시간 초과'):
            self.query(app)
        self.assertEqual(app._wait_for_erp_query_attempt.call_count, 1)

    def test_driver_failure_is_not_converted_to_no_data(self):
        app = self.app([RuntimeError('driver disconnected')])
        with self.assertRaisesRegex(RuntimeError, '조회 실패: RuntimeError'):
            self.query(app)

    def test_stop_during_retry_delay_never_clicks_again(self):
        app = self.app([{'status': 'EMPTY', 'rows': 0}])
        def stopped(_):
            app.is_running = False
            return False
        app._sleep_interruptible = stopped
        self.assertIsNone(self.query(app))
        self.assertEqual(app._wait_for_erp_query_attempt.call_count, 1)

    def attempt_app(self, snapshots):
        app = self.app([])
        app.driver.execute_script.side_effect = [None, None] + snapshots + [None]
        return app

    def attempt(self, app):
        with mock.patch('gui.time.monotonic', side_effect=lambda: self.clock):
            return RpaGuiApp._wait_for_erp_query_attempt(
                app, {'search_button': '#gridMain_r'}, '2026-11-15',
                '2026-11-15', 'LJ', '김트랑', 0.9, 0.1,
            )

    @staticmethod
    def snapshot(rows, done=True, status=200, binds=1):
        return {'requests': [{'done': done, 'status': status}],
                'binds': binds, 'boundRows': len(rows), 'rows': rows}

    def test_same_date_existing_rows_cannot_pass_pending_request(self):
        rows = [{'startDay': '20261115', 'air2Cd': 'LJ', 'priceDesc': '김트랑'}]
        pending = self.snapshot(rows, done=False, status=None, binds=0)
        app = self.attempt_app([pending] * 10)
        result = self.attempt(app)
        self.assertEqual(result['status'], 'TIMEOUT')
        self.assertIn('cleanup', app.driver.execute_script.call_args.args[0])
        app.driver.find_element.assert_called_once()

    def test_completed_empty_without_new_binding_is_not_empty(self):
        app = self.attempt_app([self.snapshot([], binds=0)] * 10)
        result = self.attempt(app)
        self.assertEqual(result['status'], 'TIMEOUT')
        self.assertIn('그리드 반영', result['reason'])

    def test_bound_completed_empty_is_empty_after_stability_wait(self):
        app = self.attempt_app([self.snapshot([])] * 10)
        result = self.attempt(app)
        self.assertEqual(result['status'], 'EMPTY')
        self.assertGreaterEqual(self.clock, 0.5)

    def test_wrong_date_airline_or_price_never_matches(self):
        for row in [
            {'startDay': '20261114', 'air2Cd': 'LJ', 'priceDesc': '김트랑'},
            {'startDay': '20261115', 'air2Cd': 'KE', 'priceDesc': '김트랑'},
            {'startDay': '20261115', 'air2Cd': 'LJ', 'priceDesc': '다낭'},
            {'startDay': '2026111', 'air2Cd': 'LJ', 'priceDesc': '김트랑'},
        ]:
            with self.subTest(row=row):
                app = self.attempt_app([self.snapshot([row])] * 10)
                self.assertEqual(self.attempt(app)['status'], 'TIMEOUT')

    def test_http_error_is_failure_and_observer_is_cleaned(self):
        app = self.attempt_app([self.snapshot([], status=500, binds=0)])
        result = self.attempt(app)
        self.assertEqual(result['status'], 'ERROR')
        self.assertIn('HTTP 500', result['reason'])
        self.assertIn('cleanup', app.driver.execute_script.call_args.args[0])

    def test_multiple_requests_fail_closed(self):
        snapshot = self.snapshot([])
        snapshot['requests'] *= 2
        app = self.attempt_app([snapshot])
        self.assertEqual(self.attempt(app)['status'], 'ERROR')

    def test_read_error_is_not_hidden_by_cleanup_error(self):
        app = self.app([])
        app.driver.execute_script.side_effect = [None, None, ValueError('read'), RuntimeError('cleanup')]
        with self.assertRaisesRegex(ValueError, 'read'):
            self.attempt(app)
        self.assertFalse(app.is_running)

    def test_cleanup_cancellation_failure_stops_all_following_queries(self):
        app = self.app([])
        app.driver.execute_script.side_effect = [None, None] + [self.snapshot([], done=False, binds=0)] * 10
        def execute(script, *args):
            if 'window.__naeilFareQuery.cleanup()' in script:
                raise RuntimeError('cancellation unverified')
            if 'const requests = state.requests' in script:
                return self.snapshot([], done=False, binds=0)
            return None
        app.driver.execute_script.side_effect = execute
        with self.assertRaisesRegex(RuntimeError, 'cancellation unverified'):
            self.attempt(app)
        self.assertFalse(app.is_running)
        app._wait_for_erp_query_attempt.reset_mock()
        self.assertIsNone(self.query(app))
        app._wait_for_erp_query_attempt.assert_not_called()

    def test_valid_binding_accepts_price_substring(self):
        rows = [{'startDay': '2026-11-15', 'air2Cd': 'LJ', 'priceDesc': '김트랑 특가'}]
        app = self.attempt_app([self.snapshot(rows)] * 10)
        self.assertEqual(self.attempt(app)['status'], 'MATCH')

    @unittest.skipUnless(shutil.which('node'), 'Node.js unavailable')
    def test_javascript_observer_waits_for_own_response_and_restores_hooks(self):
        app = self.attempt_app([self.snapshot([])] * 10)
        self.attempt(app)
        calls = app.driver.execute_script.call_args_list
        scripts = {
            'arm': calls[0].args[0], 'click': calls[1].args[0],
            'read': calls[2].args[0], 'cleanup': calls[-1].args[0],
        }
        program = r"""
            const assert = require('node:assert/strict');
            const scripts = JSON.parse(process.argv[1]);
            global.window = global;
            let rows = [{startDay:'20261115',air2Cd:'LJ',priceDesc:'same condition'}];
            global.AUIGrid = {
                setGridData: (id,data) => {rows=data;},
                getGridData: () => rows,
            };
            global.XMLHttpRequest = class {
                constructor(){this.readyState=1;this.status=0;this.handlers={};}
                send(){}
                abort(){this.readyState=0;this.aborted=true;if(this.handlers.loadend)this.handlers.loadend();}
                addEventListener(name,fn){this.handlers[name]=fn;}
                removeEventListener(name,fn){if(this.handlers[name]===fn)delete this.handlers[name];}
            };
            const originalSend=XMLHttpRequest.prototype.send;
            const originalBind=AUIGrid.setGridData;
            const arm = new Function(scripts.arm);
            const click = new Function(scripts.click);
            const read = new Function(scripts.read);
            const cleanup = new Function(scripts.cleanup);
            arm('#gridMain');
            const unrelated = new XMLHttpRequest();unrelated.send();
            assert.equal(read('#gridMain').requests.length,0);
            const xhr = new XMLHttpRequest();
            click({click(){xhr.send(); AUIGrid.setGridData('#gridMain',[]);}});
            assert.equal(read('#gridMain').binds,0); // clearing before response is not a result
            xhr.readyState=4;xhr.status=200;
            AUIGrid.setGridData('#anotherGrid',[]);
            assert.equal(read('#gridMain').binds,0);
            AUIGrid.setGridData('#gridMain',[]);
            assert.equal(read('#gridMain').binds,1);
            assert.equal(read('#gridMain').requests[0].done,false);
            xhr.handlers.loadend();
            assert.equal(read('#gridMain').requests[0].done,true);
            assert.equal(read('#gridMain').boundRows,0);
            cleanup();
            assert.equal(XMLHttpRequest.prototype.send,originalSend);
            assert.equal(AUIGrid.setGridData,originalBind);
            assert.equal(window.__naeilFareQuery,undefined);
            assert.equal(xhr.handlers.loadend,undefined);
            // Timeout/stop cleanup cancels pending delivery before a next-date click.
            arm('#gridMain');
            const pending = new XMLHttpRequest();
            click({click(){pending.send();}});
            cleanup();
            assert.equal(pending.aborted,true);
            assert.equal(pending.readyState,0);
            assert.equal(pending.handlers.loadend,undefined);
            arm('#gridMain');
            const nextDate = new XMLHttpRequest();
            click({click(){nextDate.send();}});
            assert.equal(read('#gridMain').requests.length,1);
            assert.equal(pending.readyState,0);
            cleanup();
            // A failed abort still restores hooks and reports failure.
            arm('#gridMain');
            const stuck = new XMLHttpRequest();stuck.abort=()=>{throw new Error('abort failed');};
            click({click(){stuck.send();}});
            assert.throws(()=>cleanup(),/cancellation unverified/);
            assert.equal(XMLHttpRequest.prototype.send,originalSend);
            assert.equal(AUIGrid.setGridData,originalBind);
            assert.equal(window.__naeilFareQuery,undefined);
        """
        result = subprocess.run(
            [shutil.which('node'), '-e', program, json.dumps(scripts)],
            capture_output=True, text=True, timeout=10,
        )
        self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == '__main__':
    unittest.main()
