import unittest
import json
import shutil
import subprocess
from unittest import mock

from gui import RpaGuiApp


class GridPagingTest(unittest.TestCase):
    def test_build_row_fingerprint(self):
        self.assertEqual(RpaGuiApp._build_row_fingerprint([]), "")

        keys1 = ["34635700:1001:EV01:GD01:2026-07-15:04:u1", "34635701:1002:EV02:GD02:2026-07-15:04:u2"]
        keys2 = ["34635700:1001:EV01:GD01:2026-07-15:04:u1", "34635701:1002:EV02:GD02:2026-07-15:04:u2"]
        keys3 = ["34635702:1003:EV03:GD03:2026-07-15:04:u3"]

        fp1 = RpaGuiApp._build_row_fingerprint(keys1)
        fp2 = RpaGuiApp._build_row_fingerprint(keys2)
        fp3 = RpaGuiApp._build_row_fingerprint(keys3)

        self.assertTrue(fp1.startswith("2:"))
        self.assertEqual(fp1, fp2)
        self.assertNotEqual(fp1, fp3)

    def test_row_identity_ignores_status_and_grid_generated_uid(self):
        row = {
            'basePriceSeq': 34635700,
            'eventSeq': 1001,
            'eventCd': 'EV01',
            'goodSeq': 'GD01',
            'startDay': '2026-07-15',
            'procCd': '06',
            '_$uid': 'old-render',
        }
        reloaded_row = dict(row, procCd='04', **{'_$uid': 'new-render'})
        different_row = dict(reloaded_row, basePriceSeq=34635701)

        self.assertEqual(
            RpaGuiApp._build_row_fingerprint([row]),
            RpaGuiApp._build_row_fingerprint([reloaded_row]),
        )
        self.assertNotEqual(
            RpaGuiApp._build_row_fingerprint([row]),
            RpaGuiApp._build_row_fingerprint([different_row]),
        )

    @staticmethod
    def _navigation_app_with_clock():
        app = mock.MagicMock()
        app.is_running = True
        app.config = {'paging_search_button': '#gridMain_r'}
        app._float_config = lambda *args: 0.25
        clock = [0.0]

        def sleep(seconds):
            clock[0] += seconds
            return app.is_running

        app._sleep_interruptible = sleep
        return app, clock

    def test_retry_rejects_changed_page_number_with_unchanged_rows(self):
        app, clock = self._navigation_app_with_clock()
        reads = [0]

        def snapshot():
            reads[0] += 1
            return {
                'cur': 1 if reads[0] == 1 else 2,
                'page_rows': 500,
                'fingerprint': '500:page_1_rows',
            }

        app.get_grid_page_snapshot.side_effect = snapshot
        with mock.patch('gui.time.time', side_effect=lambda: clock[0]):
            with self.assertRaisesRegex(RuntimeError, '2페이지로 이동하지 못했습니다'):
                RpaGuiApp.navigate_to_grid_page(
                    app, {'search_date_input': '#date'}, target_page=2, timeout=2.0,
                )

        self.assertEqual(app.driver.execute_script.call_count, 3)

    def test_retry_waits_for_late_target_rows(self):
        app, clock = self._navigation_app_with_clock()
        reads = [0]

        def snapshot():
            reads[0] += 1
            return {
                'cur': 1 if reads[0] == 1 else 2,
                'page_rows': 500,
                'fingerprint': '500:page_2_rows' if clock[0] >= 3.5 else '500:page_1_rows',
            }

        app.get_grid_page_snapshot.side_effect = snapshot
        with mock.patch('gui.time.time', side_effect=lambda: clock[0]):
            result = RpaGuiApp.navigate_to_grid_page(
                app, {'search_date_input': '#date'}, target_page=2, timeout=2.0,
            )

        self.assertTrue(result)
        self.assertGreaterEqual(clock[0], 3.5)
        self.assertEqual(app.driver.execute_script.call_count, 2)

    def test_stop_during_stale_retry_does_not_navigate_again(self):
        app, clock = self._navigation_app_with_clock()
        reads = [0]

        def snapshot():
            reads[0] += 1
            return {
                'cur': 1 if reads[0] == 1 else 2,
                'page_rows': 500,
                'fingerprint': '500:page_1_rows',
            }

        def sleep_and_stop(seconds):
            clock[0] += seconds
            if clock[0] >= 2.5:
                app.is_running = False
            return app.is_running

        app._sleep_interruptible = sleep_and_stop
        app.get_grid_page_snapshot.side_effect = snapshot
        with mock.patch('gui.time.time', side_effect=lambda: clock[0]):
            result = RpaGuiApp.navigate_to_grid_page(
                app, {'search_date_input': '#date'}, target_page=2, timeout=2.0,
            )

        self.assertFalse(result)
        self.assertEqual(app.driver.execute_script.call_count, 1)

    def test_get_grid_page_snapshot_from_driver(self):
        app = mock.MagicMock()
        app.config = {'grid_id': '#gridMain'}
        app.driver = mock.MagicMock()
        app.driver.execute_script.return_value = {
            'cur': 2,
            'pageRows': 500,
            'totRows': 1200,
            'totRowsReady': True,
            'declaredTotalPages': 3,
            'nextVisible': True,
            'rowKeys': ['k1', 'k2', 'k3']
        }
        app._build_row_fingerprint = RpaGuiApp._build_row_fingerprint

        snap = RpaGuiApp.get_grid_page_snapshot(app)

        self.assertEqual(snap['cur'], 2)
        self.assertEqual(snap['page_rows'], 500)
        self.assertEqual(snap['tot_rows'], 1200)
        self.assertTrue(snap['tot_rows_ready'])
        self.assertEqual(snap['declared_total_pages'], 3)
        self.assertTrue(snap['next_visible'])
        self.assertTrue(snap['fingerprint'].startswith("3:"))
        self.assertEqual(snap['row_keys'], ['k1', 'k2', 'k3'])

    def test_get_grid_page_snapshot_handles_driver_exception(self):
        app = mock.MagicMock()
        app.config = {'grid_id': '#gridMain'}
        app.driver = mock.MagicMock()
        app.driver.execute_script.side_effect = Exception("CDP disconnected")
        app._build_row_fingerprint = RpaGuiApp._build_row_fingerprint

        snap = RpaGuiApp.get_grid_page_snapshot(app)

        self.assertEqual(snap['cur'], 1)
        self.assertEqual(snap['page_rows'], 0)
        self.assertEqual(snap['tot_rows'], 0)
        self.assertFalse(snap['next_visible'])
        self.assertEqual(snap['fingerprint'], "")

    def test_navigate_to_grid_page_rejects_stale_fingerprint_and_fails(self):
        app = mock.MagicMock()
        app.is_running = True
        app.config = {'paging_search_button': '#gridMain_r'}
        app._float_config = lambda key, default, *args: 0.01
        app._sleep_interruptible = lambda s: True
        app.driver = mock.MagicMock()
        app.find_and_switch_frame = mock.MagicMock()
        app.wait_until_grid_ready_after_save = mock.MagicMock()

        # Always returns page 1's stale fingerprint even if target_page is 2
        stale_snap = {
            'cur': 1,
            'page_rows': 500,
            'tot_rows': 1000,
            'next_visible': True,
            'fingerprint': '500:stale_fp_page_1',
            'row_keys': []
        }
        app.get_grid_page_snapshot.return_value = stale_snap

        with self.assertRaises(RuntimeError) as ctx:
            RpaGuiApp.navigate_to_grid_page(app, {"search_date_input": "#date"}, target_page=2, timeout=0.05)

        self.assertIn("2페이지로 이동하지 못했습니다", str(ctx.exception))

    def test_navigate_to_grid_page_succeeds_on_new_fingerprint_stabilization(self):
        app = mock.MagicMock()
        app.is_running = True
        app.config = {'paging_search_button': '#gridMain_r'}
        app._float_config = lambda key, default, *args: 0.01
        app._sleep_interruptible = lambda s: True
        app.driver = mock.MagicMock()
        app.find_and_switch_frame = mock.MagicMock()
        app.wait_until_grid_ready_after_save = mock.MagicMock()

        pre_snap_page1 = {
            'cur': 1,
            'page_rows': 500,
            'tot_rows': 1000,
            'next_visible': True,
            'fingerprint': '500:page_1_fp',
            'row_keys': []
        }
        post_snap_stale = {
            'cur': 2,
            'page_rows': 500,
            'tot_rows': 1000,
            'next_visible': True,
            'fingerprint': '500:page_1_fp',  # stale data at first tick
            'row_keys': []
        }
        post_snap_page2 = {
            'cur': 2,
            'page_rows': 500,
            'tot_rows': 1000,
            'next_visible': False,
            'fingerprint': '500:page_2_new_fp',  # fresh page 2 data
            'row_keys': []
        }

        app.get_grid_page_snapshot.side_effect = [
            pre_snap_page1,    # pre-snapshot
            post_snap_stale,   # tick 1: rejected due to stale fp
            post_snap_page2,   # tick 2: fresh page 2 fp (stable_count = 1)
            post_snap_page2,   # tick 3: fresh page 2 fp (stable_count = 2 -> success!)
        ]

        res = RpaGuiApp.navigate_to_grid_page(app, {"search_date_input": "#date"}, target_page=2, timeout=2.0)
        self.assertTrue(res)

    def test_navigate_to_grid_page_same_target_page_reverification(self):
        app = mock.MagicMock()
        app.is_running = True
        app.config = {'paging_search_button': '#gridMain_r'}
        app._float_config = lambda key, default, *args: 0.01
        app._sleep_interruptible = lambda s: True
        app.driver = mock.MagicMock()
        app.find_and_switch_frame = mock.MagicMock()
        app.wait_until_grid_ready_after_save = mock.MagicMock()

        snap_page2 = {
            'cur': 2,
            'page_rows': 300,
            'tot_rows': 800,
            'next_visible': False,
            'fingerprint': '300:page_2_fp',
            'row_keys': []
        }

        # When prev_cur == 2 and target_page == 2, it verifies stability without rejecting matching fp
        app.get_grid_page_snapshot.side_effect = [
            snap_page2,  # pre-snapshot (cur=2)
            snap_page2,  # poll 1
            snap_page2,  # poll 2 -> stable count 2
        ]

        res = RpaGuiApp.navigate_to_grid_page(app, {"search_date_input": "#date"}, target_page=2, timeout=2.0)
        self.assertTrue(res)

    def test_get_grid_page_state_delayed_tot_rows_polling(self):
        app = mock.MagicMock()
        app.is_running = True
        app.config = {'grid_page_size': 500}
        app._float_config = lambda key, default, *args: 0.01
        app._sleep_interruptible = lambda s: True

        snap_delayed = {
            'cur': 1,
            'page_rows': 500,
            'tot_rows': 0,  # not loaded yet
            'tot_rows_ready': False,
            'declared_total_pages': 0,
            'next_visible': True,
            'fingerprint': '500:fp1',
        }
        snap_ready = {
            'cur': 1,
            'page_rows': 500,
            'tot_rows': 1200,  # loaded
            'tot_rows_ready': True,
            'declared_total_pages': 3,
            'next_visible': True,
            'fingerprint': '500:fp1',
        }

        app.get_grid_page_snapshot.side_effect = [
            snap_delayed,
            snap_ready,
            snap_ready,
        ]

        cur, total_pages, tot, next_vis = RpaGuiApp.get_grid_page_state(app, timeout=2.0)

        self.assertEqual(cur, 1)
        self.assertEqual(total_pages, 3)  # ceil(1200 / 500) = 3
        self.assertEqual(tot, 1200)
        self.assertTrue(next_vis)

    def test_get_grid_page_state_next_visible_conservative_fallback(self):
        app = mock.MagicMock()
        app.is_running = True
        app.config = {'grid_page_size': 500}
        app._float_config = lambda key, default, *args: 0.01
        app._sleep_interruptible = lambda s: True

        # When tot_rows is 0 or low, but next_visible is True, total_pages should be at least 2
        snap_next_visible_only = {
            'cur': 1,
            'page_rows': 500,
            'tot_rows': 0,
            'tot_rows_ready': False,
            'declared_total_pages': 0,
            'next_visible': True,
            'fingerprint': '',
        }
        app.get_grid_page_snapshot.return_value = snap_next_visible_only

        cur, total_pages, tot, next_vis = RpaGuiApp.get_grid_page_state(app, timeout=0.0)

        self.assertEqual(cur, 1)
        self.assertEqual(total_pages, 2)
        self.assertTrue(next_vis)

        # Current totRows overrides stale paging controls from the previous query.
        snap_partial_with_next = {
            'cur': 1,
            'page_rows': 450,
            'tot_rows': 450,
            'tot_rows_ready': True,
            'declared_total_pages': 2,
            'next_visible': True,
            'fingerprint': '450:fp',
        }
        app.get_grid_page_snapshot.return_value = snap_partial_with_next

        cur, total_pages, tot, next_vis = RpaGuiApp.get_grid_page_state(app, timeout=0.0)

        self.assertEqual(cur, 1)
        self.assertEqual(total_pages, 1)
        self.assertTrue(next_vis)

    def test_current_total_overrides_previous_query_paging(self):
        for total, expected in [(16, 1), (40, 1), (800, 2), (1123, 3)]:
            with self.subTest(total=total):
                app = mock.MagicMock()
                app.is_running = True
                app.config = {'grid_page_size': 500}
                app.get_grid_page_snapshot.return_value = {
                    'cur': 1, 'page_rows': min(total, 500), 'tot_rows': total,
                    'tot_rows_ready': True, 'declared_total_pages': 8,
                    'next_visible': True, 'fingerprint': 'current-query',
                }
                self.assertEqual(RpaGuiApp.get_grid_page_state(app)[1], expected)

    def test_pending_row_selection_uses_model_for_offscreen_row(self):
        app = mock.MagicMock()
        app.config = {'grid_id': '#gridMain'}
        app._current_page_progress_rows.return_value = [
            {'procCd': '06' if i == 172 else '04'} for i in range(500)
        ]
        app.driver.execute_script.side_effect = [{'ok': True, 'selected': [172]}, 1]
        with mock.patch('gui.time.sleep'):
            self.assertEqual(RpaGuiApp._select_current_page_rows_by_progress_status(
                app, {}, '06', 0, 0.1,
            ), 1)
        selection_call = app.driver.execute_script.call_args_list[0]
        self.assertEqual(selection_call.args[1:], ('#gridMain', [172], '06'))
        self.assertIn('AUIGrid.setCheckedRowsByValue', selection_call.args[0])
        self.assertNotIn('document.querySelectorAll', selection_call.args[0])

    def test_pending_row_selection_fails_before_modal_on_wrong_selection(self):
        app = mock.MagicMock()
        app.config = {}
        app._current_page_progress_rows.return_value = [{'procCd': '06'}]
        app.driver.execute_script.return_value = {
            'ok': False, 'reason': 'target_rows_not_checked', 'selected': [1],
        }
        with self.assertRaisesRegex(RuntimeError, 'target_rows_not_checked'):
            RpaGuiApp._select_current_page_rows_by_progress_status(app, {}, '06', 0, 0.1)

    @unittest.skipUnless(shutil.which('node'), 'Node.js is required for grid JavaScript regression')
    def test_pending_selection_javascript_virtual_rows_and_disabled_policy(self):
        app = mock.MagicMock()
        app.config = {}
        app._current_page_progress_rows.return_value = [{'procCd': '06'}]
        app.driver.execute_script.return_value = {'ok': False, 'reason': 'capture'}
        with self.assertRaises(RuntimeError):
            RpaGuiApp._select_current_page_rows_by_progress_status(app, {}, '06', 0, 0.1)
        script = app.driver.execute_script.call_args.args[0]
        harness = """
        const select = new Function(SCRIPT);
        const rows = Array.from({length:500}, (_,i)=>({procCd:i===172?'06':'04'}));
        let checked = [{rowIndex:0,item:rows[0]}], blocked = false, calls = 0;
        global.AUIGrid = {
          getGridData:()=>rows,
          getCheckedRowItems:()=>checked,
          getProp:(_,name)=>name==='rowCheckDisabledFunction'?(()=>!blocked):null,
          setCheckedRowsByValue:(_,field,value)=>{
            calls++; checked=rows.flatMap((item,rowIndex)=>item[field]===value?[{rowIndex,item}]:[]);
          }
        };
        const first=select('#gridMain',[172],'06');
        if(!first.ok || first.selected[0]!==172 || checked.length!==1) throw Error('offscreen selection failed');
        blocked=true;
        const second=select('#gridMain',[172],'06');
        if(second.ok || second.reason!=='target_row_not_selectable' || calls!==1) throw Error('disabled policy bypassed');
        blocked=false; AUIGrid.setCheckedRowsByValue=()=>{checked=[{rowIndex:0,item:rows[0]}]};
        if(select('#gridMain',[172],'06').ok) throw Error('incorrect selected row accepted');
        """.replace('SCRIPT', json.dumps(script))
        result = subprocess.run([shutil.which('node'), '-e', harness], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_get_grid_page_state_fails_closed_when_page_count_never_arrives(self):
        app = mock.MagicMock()
        app.is_running = True
        app.config = {'grid_page_size': 500}
        app._float_config = lambda key, default, *args: 0.01
        app._sleep_interruptible = lambda s: True
        app.get_grid_page_snapshot.return_value = {
            'cur': 1,
            'page_rows': 500,
            'tot_rows': 0,
            'tot_rows_ready': False,
            'declared_total_pages': 0,
            'next_visible': True,
            'fingerprint': '500:fp',
        }

        with self.assertRaisesRegex(RuntimeError, '전체 페이지 수를 확인하지 못했습니다'):
            RpaGuiApp.get_grid_page_state(app, timeout=0.02, poll_interval=0.01)

    def test_get_grid_page_state_ignores_stale_larger_declared_page_count(self):
        app = mock.MagicMock()
        app.is_running = True
        app.config = {'grid_page_size': 500}
        app._float_config = lambda key, default, *args: 0.01
        app._sleep_interruptible = lambda s: True
        stable_snapshot = {
            'cur': 1,
            'page_rows': 500,
            'tot_rows': 500,
            'tot_rows_ready': True,
            'declared_total_pages': 3,
            'next_visible': True,
            'fingerprint': '500:stable',
        }
        app.get_grid_page_snapshot.side_effect = [stable_snapshot, stable_snapshot]

        cur, total_pages, tot, next_vis = RpaGuiApp.get_grid_page_state(
            app,
            timeout=1.0,
            poll_interval=0.01,
        )

        self.assertEqual(cur, 1)
        self.assertEqual(total_pages, 1)
        self.assertEqual(tot, 500)
        self.assertTrue(next_vis)

    def test_get_grid_page_state_fails_closed_when_snapshot_never_stabilizes(self):
        app = mock.MagicMock()
        app.is_running = True
        app.config = {'grid_page_size': 500}
        app._float_config = lambda key, default, *args: 0.01
        app._sleep_interruptible = lambda s: True
        fingerprints = iter(range(1000000))

        def changing_snapshot():
            return {
                'cur': 1,
                'page_rows': 500,
                'tot_rows': 1500,
                'tot_rows_ready': True,
                'declared_total_pages': 3,
                'next_visible': True,
                'fingerprint': f"500:changing-{next(fingerprints)}",
            }

        app.get_grid_page_snapshot.side_effect = changing_snapshot

        with self.assertRaisesRegex(RuntimeError, '페이지 정보가 안정화되지 않았습니다'):
            RpaGuiApp.get_grid_page_state(app, timeout=0.02, poll_interval=0.01)


if __name__ == '__main__':
    unittest.main()
