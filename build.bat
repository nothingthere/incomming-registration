pyinstaller --noconfirm ^
            --noconsole ^
            --clean ^
            --add-data="sdgs.png;." ^
            --add-data="sdgs.ico;." ^
            --add-data="folder.png;." ^
            --add-data="excel.png;." ^
            -i banana.ico ^
            --name="receiving-registor" ^
            root.py
