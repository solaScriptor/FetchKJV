pyinstaller --onefile ^
  --icon=FetchKJV.ico ^
  --add-data "kjv.json;." ^
  --add-data "settings.json;." ^
  --add-data "FetchKJV.ico;." ^
  --add-data "FetchKJV.png;." ^
  --noconsole ^
  FetchKJV.py
