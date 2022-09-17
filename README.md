# GFLBattleSim-JSONGenerator
<img src="https://i.imgur.com/GLtrdD2.png" width="100%">

# General Overview

This is a macro-enabled Excel spreadsheet that will generate/modify the JSON files needed for the GFL KR Battle Sim.

Among other things, this spreadsheet can be used to more easily:

* Change battle settings (day/night/target practice, enemy ID, boss HP, HOC/Fairy, etc.)
* Edit G&K echelon composition (dolls, skill levels, equipment, positions, etc.)
* Edit SF echelon composition (units, skill levels, chips, positions, etc.)
* Edit fairy information (fairy, rarity, level, skill level, etc.)
* Edit HOC and SF HOC information (stats, skill level, etc.)
* Apply strategy fairy skills (Parachute, Construction, Suee, Combo, etc.)
* Apply enemy debuffs (currently only EMP, Auspicious)
* Load/save/delete preset echelon

VBA code is provided in a separate .vb file so it can be examined. Place the spreadsheet in the parent directory of the battle sim (i.e., the same folder as GFBattleSimulator.exe).

# IMPORTANT: Client v3.01 and EN Compatibility

The [original battle sim](https://gall.dcinside.com/mgallery/board/view?id=micateam&no=1506585) by kchang will mostly work on client v3.01 if using the userinfo.json from this repo, but an error will make TargetTrain mode inaccessible. kchang has kindly provided me with a copy of the source code and I have modified it so that TargetTrain mode works and so that the sim works with the EN client; however, there are some OS-specific bugs for EN that I have not fully figured out yet, so please ues the following workarounds in the meantime:

Android EN (tested with blue MuMu):
1. Start GFL while connected to the sim's proxy
2. Get to the login screen and try to log in (it will hang up and fail)
3. Disconnect the proxy and try to log in again (it should now log you onto the sim account)
4. Reconnect the proxy after logging in and the sim should work normally
5. Redo the above if you restart client. If you just want to reload and update sim settings, trigger a Code3 instead of restarting client (end turn, terminate map, change real echelon’s formation, remove equip, etc.)

iOS EN:
1. Start GFL while using cellular data or while not connected to proxy
2. At the login screen, press the “Switch Account” button
3. Connect to proxy and play as guest. Sim should work normally from there.
4. Need to redo the above even if you trigger a Code3.

Modified sim download link: https://drive.google.com/file/d/1RkKi0jXYPn8r4qAmo8aDBPc7BvcptVcz/view?usp=sharing

# Do I Need to Use Excel?
This spreadsheet was created with Excel 2016 and I highly recommended using Excel to avoid any unforeseen bugs or issues. While I cannot guarantee that the spreadsheet is fully compatible with older versions, I have done my best to avoid using formulas or VBA that may not be available on older versions.

Google Sheets does not support VBA and cannot be used, but I have tested [LibreOffice v7.3.2](https://www.libreoffice.org/) (the latest version at the time of this commit) and have adjusted the spreadsheet so that its core functionality (e.g., using the dropdowns to generate JSONs) is intact. Known issues with LibreOffice that I will likely not fix because they are VBA-related:
* Loading echelon presets will work, but saving presets throws an error and deleting presets breaks the ability to load.
* The "Download Latest Version from GitHub" button does not work.

# Additional Usage Notes:
* As mentioned above, please replace the default userinfo.json with the one in this repo.
* When entering G&K HOC stats, use the FINAL stats including chips and iteration bonuses. The base stats will be adjusted to get the desired stats, so your HOCs will appear to be max iteration and will have blank chip boards.
* You cannot sim more than 8 G&K HOCs at a time. This appears to be an issue with the sim itself.
* There are issues with simming G&K directly after simming SF echelons. If you want to switch back to G&K, you should generate the new JSONs (or use the ones from the original sim download) and restart the sim .exe.
* To properly sim Desert Fairy, you should activate both components in the strategy skill section in this order: "Desert 1 (negation)" and "Desert 2 (debuff)". The  latter applies the debuff to you and the enemy, the former negates the effect on your echelon.

# To Do:
* Possible code/formatting clean up

# Other Links:
* KR Battle Sim: https://gall.dcinside.com/mgallery/board/view?id=micateam&no=1506585
* GamePress guide: https://gamepress.gg/girlsfrontline/how-use-gf-battle-tester-girls-frontline-battle-tester
* EN Battle Sim: https://github.com/neko-gg/gfl-combat-simulator
* Data Source: https://github.com/randomqwerty/GFLData

# Contact Info:
* Discord: Randomqwerty#4678
* Reddit: /u/UnironicWeeaboo
