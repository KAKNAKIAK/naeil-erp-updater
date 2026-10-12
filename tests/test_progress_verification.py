import unittest
import json
import shutil
import subprocess
from unittest import mock

from gui import RpaGuiApp


class ProgressVerificationTest(unittest.TestCase):
    def test_every_progress_code_requires_every_target_row(self):
        for code in ['01', '02', '04', '05', '06']:
            with self.subTest(code=code):
                rows = [{'basePriceSeq': 1, 'procCd': code}, {'basePriceSeq': 2, 'procCd': code}]
                RpaGuiApp._verify_progress_rows(rows, {'1', '2'}, code)
                rows[1]['procCd'] = 'different'
                with self.assertRaisesRegex(RuntimeError, '상태 불일치 1건'):
                    RpaGuiApp._verify_progress_rows(rows, {'1', '2'}, code)

    def test_missing_saved_rows_are_failure(self):
        with self.assertRaisesRegex(RuntimeError, '누락 1건'):
            RpaGuiApp._verify_progress_rows([{'basePriceSeq': 1, 'procCd': '05'}], {'1', '2'}, '05')

    def test_identity_missing_zero_and_duplicates_fail(self):
        for rows in [[], [{}], [{'basePriceSeq': 0}], [{'basePriceSeq': 1}, {'basePriceSeq': '1'}]]:
            with self.subTest(rows=rows), self.assertRaises(RuntimeError):
                RpaGuiApp._progress_row_ids(rows)

    def test_hotel_subset_verification_does_not_expand_scope(self):
        RpaGuiApp._verify_progress_rows([
            {'basePriceSeq': 1, 'procCd': '04'}, {'basePriceSeq': 2, 'procCd': '05'},
        ], {'1'}, '04')

    def test_page_response_failure_cannot_pass_unchanged_snapshot(self):
        app = mock.MagicMock()
        app.is_running = True
        app.config = {}
        app._float_config.return_value = .1
        app.get_grid_page_snapshot.return_value = {'cur': 1, 'fingerprint': 'same', 'page_rows': 2}
        app._wait_for_erp_query_attempt.return_value = {'status': 'TIMEOUT'}
        with self.assertRaisesRegex(RuntimeError, '재조회 확인 실패'):
            RpaGuiApp.navigate_to_grid_page(app, {'search_date_input': '#date'}, 1, 2)

    def test_all_selection_propagates_grid_selection_failure(self):
        app = mock.MagicMock()
        app.config = {}
        app._current_page_progress_rows.return_value = [{'procCd': '04'}, {'procCd': '05'}]
        app.driver.execute_script.return_value = {'ok': False, 'reason': 'target_row_not_selectable'}
        with self.assertRaisesRegex(RuntimeError, 'target_row_not_selectable'):
            RpaGuiApp._select_current_page_rows_by_progress_status(app, {}, None, 0, .1)
        self.assertEqual(app.driver.execute_script.call_args.args[2:], ([0, 1], None, ''))

    @unittest.skipUnless(shutil.which('node'), 'Node.js required')
    def test_all_selection_erp_excludes_goal_but_cannot_omit_non_goal(self):
        app = mock.MagicMock()
        app.config = {}
        app._current_page_progress_rows.return_value = [{'procCd': '04'}]
        app.driver.execute_script.return_value = {'ok': False, 'reason': 'capture'}
        with self.assertRaises(RuntimeError):
            RpaGuiApp._select_current_page_rows_by_progress_status(app, {}, None, 0, .1, '05')
        script = app.driver.execute_script.call_args.args[0]
        harness = '''
        const select = new Function(SCRIPT);
        const rows = Array.from({length:500}, (_,i)=>({basePriceSeq:i%2?String(i+1):i+1,procCd:i<10?'05':'04'}));
        let checked=[], omit=false;
        global.AUIGrid={
          getGridData:()=>rows, getCheckedRowItems:()=>checked,
          getProp:(_,name)=>name==='rowCheckDisabledFunction'?((i,c,r)=>r.procCd!=='05'):
            name==='independentAllCheckBox'?true:null,
          setCheckedRowsByValue:()=>{throw Error('general selection used subset API')},
          setAllCheckedRows:()=>{},
          setCheckedRowsByIds:(_,ids)=>{checked=rows.flatMap((item,rowIndex)=>
            ids.includes(item.basePriceSeq) && !(omit && rowIndex===100)?[{rowIndex,item}]:[])}
        };
        const indexes=rows.map((_,i)=>i);
        const result=select('#gridMain',indexes,null,'05');
        if(!result.ok || result.selected.length!==490) throw Error('ERP exclusion rejected');
        if(select('#gridMain',indexes,null,'04').ok) throw Error('closed non-goal exclusion accepted');
        omit=true;
        if(select('#gridMain',indexes,null,'05').ok) throw Error('missing non-goal accepted');
        rows.forEach(row=>row.procCd='05'); omit=false;
        if(!select('#gridMain',indexes,null,'05').ok) throw Error('all-goal page rejected');
        checked=[{rowIndex:500,item:{procCd:'04'}}];
        AUIGrid.setCheckedRowsByIds=()=>{};
        if(select('#gridMain',indexes,null,'05').ok) throw Error('unexpected selection accepted');
        '''.replace('SCRIPT', json.dumps(script))
        result = subprocess.run([shutil.which('node'), '-e', harness], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_all_goal_page_requeries_and_verifies_without_save(self):
        app = mock.MagicMock()
        app.config = {}
        rows = [{'basePriceSeq': 1, 'procCd': '05'}]
        app._current_page_progress_rows.side_effect = [rows, rows]
        app._progress_row_ids.side_effect = RpaGuiApp._progress_row_ids
        app._verify_progress_rows.side_effect = RpaGuiApp._verify_progress_rows
        app._select_current_page_rows_by_progress_status.return_value = 0
        app.driver.execute_script.return_value = []
        self.assertEqual(RpaGuiApp._apply_current_page_progress_status(
            app, {}, '05', 'closed', 2, 0, .1), 0)
        app.navigate_to_grid_page.assert_called_once_with({}, 1, 2)
        app.driver.find_element.assert_not_called()

    def test_missing_non_goal_selection_stops_before_popup(self):
        app = mock.MagicMock()
        app.config = {}
        app._current_page_progress_rows.return_value = [
            {'basePriceSeq': 1, 'procCd': '05'}, {'basePriceSeq': 2, 'procCd': '04'}]
        app._progress_row_ids.side_effect = RpaGuiApp._progress_row_ids
        app._select_current_page_rows_by_progress_status.return_value = 0
        app.driver.execute_script.return_value = []
        with self.assertRaisesRegex(RuntimeError, '선택 행 ID'):
            RpaGuiApp._apply_current_page_progress_status(app, {}, '05', 'closed', 2, 0, .1)
        app.driver.find_element.assert_not_called()

    def test_mixed_page_still_verifies_excluded_goal_row_after_save(self):
        app = mock.MagicMock()
        app.config = {}
        before = [{'basePriceSeq': 1, 'procCd': '05'}, {'basePriceSeq': 2, 'procCd': '04'}]
        after = [{'basePriceSeq': 2, 'procCd': '05'}]
        app._current_page_progress_rows.side_effect = [before, after]
        app._progress_row_ids.side_effect = RpaGuiApp._progress_row_ids
        app._verify_progress_rows.side_effect = RpaGuiApp._verify_progress_rows
        app._select_current_page_rows_by_progress_status.return_value = 1
        app.driver.execute_script.side_effect = [[before[1]], None,
            {'ok': True, 'checked': True, 'value': '05'}, None]
        app.driver.switch_to.alert.text = 'saved'
        with mock.patch('gui.WebDriverWait'), mock.patch('gui.time.sleep'):
            with self.assertRaisesRegex(RuntimeError, '누락 1건'):
                RpaGuiApp._apply_current_page_progress_status(app, {}, '05', 'closed', 2, 0, .1)

    def test_closed_modal_does_not_hide_missing_saved_status(self):
        app = mock.MagicMock()
        app.config = {}
        before = [{'basePriceSeq': 1, 'procCd': '04'}, {'basePriceSeq': 2, 'procCd': '04'}]
        after = [{'basePriceSeq': 1, 'procCd': '05'}, {'basePriceSeq': 2, 'procCd': '04'}]
        app._current_page_progress_rows.side_effect = [before, after]
        app._progress_row_ids.side_effect = RpaGuiApp._progress_row_ids
        app._verify_progress_rows.side_effect = RpaGuiApp._verify_progress_rows
        app._select_current_page_rows_by_progress_status.return_value = 2
        app.navigate_to_grid_page.return_value = True
        app.driver.execute_script.side_effect = [before, None,
            {'ok': True, 'checked': True, 'value': '05'}, None]
        app.driver.switch_to.alert.text = '저장 완료'
        with mock.patch('gui.WebDriverWait'), mock.patch('gui.time.sleep'):
            with self.assertRaisesRegex(RuntimeError, '상태 불일치 1건'):
                RpaGuiApp._apply_current_page_progress_status(
                    app, {}, '05', '예약마감', 2, 0, .1, target_page=2)
        app.navigate_to_grid_page.assert_called_once_with({}, 2, 2)

    def test_changed_selection_ids_stop_before_save_popup(self):
        app = mock.MagicMock()
        app.config = {}
        app._current_page_progress_rows.return_value = [{'basePriceSeq': 1, 'procCd': '04'}]
        app._progress_row_ids.side_effect = RpaGuiApp._progress_row_ids
        app._select_current_page_rows_by_progress_status.return_value = 1
        app.driver.execute_script.return_value = [{'basePriceSeq': 9, 'procCd': '04'}]
        with self.assertRaisesRegex(RuntimeError, '선택 행 ID'):
            RpaGuiApp._apply_current_page_progress_status(app, {}, '05', '예약마감', 2, 0, .1)
        app.driver.find_element.assert_not_called()


if __name__ == '__main__':
    unittest.main()
