# -*- coding: utf-8 -*-
import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def main():
    print("디버깅 크롬 포트 9222 연결 시도...")
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print(f"드라이버 연결 실패: {e}")
        return
        
    try:
        driver.switch_to.default_content()
        iframe = driver.find_element(By.ID, "framego01_104_list")
        driver.switch_to.frame(iframe)
        print("framego01_104_list iframe으로 스위칭 완료.")
        
        st_date_input = driver.find_element(By.CSS_SELECTOR, "#searchStDate")
        en_date_input = driver.find_element(By.CSS_SELECTOR, "#searchEnDate")
        
        # 날짜 임의 변경 테스트
        target_date = "2026-10-12"
        print(f"날짜 변경 시도: {target_date}")
        driver.execute_script("arguments[0].value = '';", st_date_input)
        st_date_input.send_keys(target_date)
        
        driver.execute_script("arguments[0].value = '';", en_date_input)
        en_date_input.send_keys(target_date)
        
        print(f"현재 입력 필드 확인 -> 시작일: '{st_date_input.get_attribute('value')}', 종료일: '{en_date_input.get_attribute('value')}'")
        
        # 1차 시도: 부모 a 태그 클릭 시도
        search_btns = driver.find_elements(By.CSS_SELECTOR, "a:has(#gridMain_r), #gridMain_r")
        print(f"총 {len(search_btns)}개의 조회 버튼 감지")
        
        # a 태그 클릭 시도
        a_btn = search_btns[0]
        print(f"a 태그 클릭 시도... (태그: {a_btn.tag_name})")
        driver.execute_script("arguments[0].click();", a_btn)
        
        time.sleep(3.0)
        driver.save_screenshot("screenshot_after_a_click.png")
        print("a 태그 클릭 후 스크린샷 저장 완료: screenshot_after_a_click.png")
        
        # 2차 시도: 이미지 태그 직접 클릭 시도 (#gridMain_r)
        img_btn = driver.find_element(By.ID, "gridMain_r")
        print("img#gridMain_r 직접 클릭 시도...")
        driver.execute_script("arguments[0].click();", img_btn)
        
        time.sleep(3.0)
        driver.save_screenshot("screenshot_after_img_click.png")
        print("img 클릭 후 스크린샷 저장 완료: screenshot_after_img_click.png")
        
    except Exception as e:
        print(f"디버깅 중 에러 발생: {e}")
        
    driver.quit()

if __name__ == "__main__":
    main()
