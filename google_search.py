import traceback
import uiautomation as auto
from datetime import datetime
import time
import logging

logger = logging.getLogger(__file__)
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s %(levelname)-.1s %(lineno)4s:%(funcName)20s %(message)s')
handler = logging.StreamHandler()
handler.setFormatter(formatter)
logger.addHandler(handler)


def captcha_check(browser):
    captcha = browser.CheckBoxControl(Name="reCAPTCHA 인증 필요. 로봇이 아닙니다.")
    if captcha.Exists(2, 1):
        captcha.Click()
        time.sleep(3)
        home = browser.ButtonControl(Name="홈")
        if home.Exists(2, 1):
            home.Click()
            return True


def chrome_google(keyword, target):
    start_fun = datetime.now()
    result = False

    try:
        logger.info("="*80)
        logger.info(f"keyword: {keyword} / target: {target}")
        logger.info("=" * 80)

        logger.info(f"[1/6] 크롬 브라우저 검색")
        browser = auto.PaneControl(ClassName="Chrome_WidgetWin_1")
        if browser.Exists(5, 1):
            browser.SetActive()
        else:
            return False

        logger.info(f"[2/6] 글꼴 크기 100%로 변경")
        auto.SendKeys("{Ctrl}0")

        if captcha_check(browser):
            return False

        logger.info(f"[3/6] URL 입력 박스에 URL 입력")
        control = browser.EditControl(Name="주소창 및 검색창")
        if control.Exists(5, 1):
            control.Click()
            control.SendKeys("https://google.co.kr")
            control.SendKeys("{Enter}")
            time.sleep(1)
        else:
            return False

        if captcha_check(browser):
            return False

        logger.info(f"[4/6] 페이지 로드 확인")
        refresh = browser.ButtonControl(Name="새로고침")
        while refresh.GetPropertyValue(30094) == '해당 페이지 로드 중지':
            time.sleep(0.5)
            logger.info(f" - {refresh.GetPropertyValue(30094)}'")

        if captcha_check(browser):
            return False

        logger.info(f"[5/6] 검색 키워드 입력")
        control = browser.ComboBoxControl(Name="검색")
        if control.Exists(5, 1):
            while control.BoundingRectangle.width() == 0 and control.BoundingRectangle.height() == 0:
                logger.info(f' - 클릭 가능할때까지 1초 대기')
                time.sleep(1)
            control.Click()
            control.SendKeys(keyword, interval=0.1)
            control.SendKeys("{Enter}")
        else:
            return False

        if captcha_check(browser):
            return False

        logger.info(f"[6/6] 검색 결과 페이지에서 제목 검색")
        for page_no in range(1, 10 + 1):

            if captcha_check(browser):
                return False

            logger.info(f' - {page_no} 페이지 검색')
            if page_no == 1:
                pass
            else:
                control = browser.HyperlinkControl(Name="다음")
                if control.Exists(5, 1):
                    auto.SendKey(auto.Keys.VK_END)
                    control.Click()
                else:
                    return False

            control = browser.HyperlinkControl(SubName=target)

            if control.Exists(2, 0.5):

                if control.BoundingRectangle.width() == 0 and control.BoundingRectangle.height() == 0:
                    auto.SendKey(auto.Keys.VK_PAGEDOWN)
                else:
                    auto.SendKey(auto.Keys.VK_DOWN)

                if type(control) == auto.uiautomation.HyperlinkControl:
                    control.Click()
                else:
                    text_control = control.GetParentControl().GetPreviousSiblingControl().GetLastChildControl()
                    link_control = text_control.GetFirstChildControl()
                    link_control.Click()
                time.sleep(1)

                tab = auto.DocumentControl(Value=target)
                if tab.Exists(10, 1):
                    time.sleep(3)

                    for no in range(1, 6):
                        auto.SendKey(auto.Keys.VK_PAGEDOWN)
                        logger.info(f" - [{no}/{5}] PAGEDOWN 키 누름")
                        time.sleep(0.5)

                    result = True
                    break
                else:
                    break

    except Exception as ex:
        logger.error(f"{traceback.format_exc()}")

    finally:
        logger.info(f"소요시간 {(datetime.now() - start_fun).total_seconds():.3f}")

    return result


if __name__ == '__main__':
    chrome_google('1등로또필터 당첨', 'https://mobile.newsis.com')
