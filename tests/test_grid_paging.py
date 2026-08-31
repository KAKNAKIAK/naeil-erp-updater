import unittest
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

        # When tot_rows <= 500 (e.g. 450) but next_visible is True, total_pages should be at least 2
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
        self.assertEqual(total_pages, 2)
        self.assertTrue(next_vis)

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

    def test_get_grid_page_state_uses_larger_declared_page_count(self):
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
        self.assertEqual(total_pages, 3)
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
