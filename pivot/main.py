import json
import gspread
from sheets import sheet1

CONFIG_PATH = "config.json"

SHEET_MODULES = {
    "sheet1": sheet1,
}


def main():
    with open(CONFIG_PATH) as f:
        config = json.load(f)

    gc = gspread.oauth()

    for sheet_cfg in config["sheets"]:
        name = sheet_cfg["name"]
        url = sheet_cfg["url"]
        module = SHEET_MODULES.get(name)
        if module is None:
            print(f"  {name}: 모듈 없음, 건너뜀")
            continue
        print(f"처리 중: {name}")
        sh = gc.open_by_url(url)
        module.run(sh)

    print("완료")


if __name__ == "__main__":
    main()
