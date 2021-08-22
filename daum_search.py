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


def chrome_daum(keyword, target):
    start_fun = datetime.now()
    result = False

    try:
        logger.info("=" * 80)
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

        logger.info(f"[3/6] URL 입력 박스에 URL 입력")
        control = browser.EditControl(Name="주소창 및 검색창")
        if control.Exists(5, 1):
            control.Click()
            control.SendKeys("https://search.daum.net/search?w=web&nil_search=btn&DA=NTB&enc=utf8&lpp=10&q=")
            control.SendKeys("{Enter}")
        else:
            return False

        logger.info(f"[4/6] 페이지 로드 확인")
        while browser.Name != "– Daum 검색 - Chrome":
            time.sleep(0.1)

        browser = auto.PaneControl(Name="– Daum 검색 - Chrome")
        if browser.Exists(20, 1):
            pass
        else:
            return False

        time.sleep(1)

        logger.info(f"[5/6] 검색 키워드 입력")
        control = browser.EditControl(Name="검색어 입력")
        if control.Exists(5, 1):
            while control.BoundingRectangle.width() == 0 and control.BoundingRectangle.height() == 0:
                time.sleep(1)
            control.Click()
            control.SendKeys(keyword, interval=0.1)
            control.SendKeys("{Enter}")
        else:
            return False

        logger.info(f"[6/6] 검색 결과 페이지에서 제목 검색")
        browser = auto.PaneControl(Name=f"{keyword} – Daum 검색 - Chrome")
        for page_no in range(1, 10 + 1):

            logger.info(f' - {page_no} 페이지 검색')
            if page_no == 1:
                pass
            else:
                control = browser.HyperlinkControl(Name=f"{page_no}")
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

                control.Click()
                time.sleep(1)

                tab = auto.PaneControl(SubName=target)
                if tab.Exists(10, 1):

                    time.sleep(1)
                    tab.SendKeys("{Tab}")
                    time.sleep(3)

                    for no in range(1, 8):
                        auto.SendKey(auto.Keys.VK_PAGEDOWN)
                        logger.info(f"[{no}/{7}] PAGEDOWN 키 누름")
                        time.sleep(0.5)

                    result = True
                    break
                else:
                    break

    except Exception as e:
        logger.error(f"{traceback.format_exc()}")

    finally:
        logger.info(f"소요시간 {(datetime.now() - start_fun).total_seconds():.3f}")

    return result


if __name__ == '__main__':
    chrome_daum('로또블루', '20대가 로또되면')
