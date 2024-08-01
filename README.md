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
* Edit Mobile Armor information (tech tree, crew, components, combat modes, etc.)
* Apply strategy fairy skills (Parachute, Construction, Suee, Combo, etc.)
* Apply enemy debuffs (currently only EMP, Auspicious)
* Load/save/delete preset echelon

VBA code is provided in a separate .vb file so it can be examined. Place the spreadsheet in the parent directory of the battle sim (i.e., the same folder as GFBattleSimulator.exe).

# IMPORTANT: Client v3.01+ and EN Compatibility

Note: As of December 14, 2023, you will need to use the files in the "Pre-3.03" folder if you want to want to use this for clients older than v3.03. The same .exe can be used for pre/post-v3.03.

The [original battle sim](https://gall.dcinside.com/mgallery/board/view?id=micateam&no=1506585) made by kchang had issues with v3.01 and will no longer work as of v3.02, due to new responses that the server expects. kchang has kindly provided me with a copy of the source code and I have modified it to fake the responses needed and so that the sim works with the EN client; however, there are some OS-specific bugs for EN that I have not fully figured out yet, so please use the following workarounds in the meantime:

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

Modified sim download link: https://drive.google.com/file/d/1VOceIplePJDXXPI6NZja62NK5Yor2Kv2/view?usp=sharing

This sim has been modified to let it read custom responses from a new `preset\responses\` folder, similar to GFAlarm's integration of the battle sim that is buggy with EN. This should hopefully allow users to fix any new errors that appear when a client update occurs.

# Theater Testing

For now, this will likely only work while Theater is live and you will not be able to challenge zones Elem/Int/Adv once the server enters Core. It is also not an automated process. I may see if this is possible to change later and I may also need to update the instructions below once we get another Theater and I can test more.

1. Open GFAlarm. Click the gear icon to go into settings, enable "Print Packet Log" (under the Extras section)
2. Log into GFL and open the Theater menu. You can turn "Print Packet Log" off now.
3. Go to the `Log` subfolder of your GFAlarm folder, open the latest log, look for the one that says `"URL":"/Theater/data"`.
4. Copy the part after `"Response":` that is enclosed in braces {} and make sure that it is a valid JSON (delete the last right brace that is leftover)
5. Paste that into `preset\theater_data.json` in your Battle Sim folder.
6. C;ear anything inside `theater_teams_info`. It should look like this: `"theater_teams_info":{},`
7. Set the sim to TargetTrain mode, generate your echelon/equipment/fairy/etc., log in, enter Theater, select your map, and play as usual.

Note that you may be able to skip steps 1-5 by updating the file manually, but that needs more testing to make sure it works. If you do not want to clear all waves, you can also go into `preset\responses\startTheaterExercise` and remove waves from the list (it is zero-indexed, so first wave = 0, last wave = 5).

# Do I Need to Use Excel?
This spreadsheet was created with Excel 2016 and I highly recommended using Excel to avoid any unforeseen bugs or issues. While I cannot guarantee that the spreadsheet is fully compatible with older versions, I have done my best to avoid using formulas or VBA that may not be available on older versions.

Google Sheets does not support VBA and cannot be used, but I have tested [LibreOffice v7.6.2](https://www.libreoffice.org/) (the latest version at the time of this commit) and have adjusted the spreadsheet so that its core functionality (e.g., using the dropdowns to generate JSONs) is intact. Known issues with LibreOffice that I will likely not fix because they are VBA-related:
* Loading echelon presets will work, but saving presets throws an error and deleting presets breaks the ability to load.
* The "Download Latest Version from GitHub" button does not work.

# Additional Usage Notes:
* As mentioned above, please replace the default userinfo.json with the one in this repo.
* When entering G&K HOC stats, use the FINAL stats including chips and iteration bonuses. The base stats will be adjusted to get the desired stats, so your HOCs will appear to be max iteration and will have blank chip boards.
* You cannot sim more than 8 G&K HOCs at a time. This appears to be an issue with the sim itself.
* There are issues with simming G&K directly after simming SF echelons. If you want to switch back to G&K, you should generate the new JSONs (or use the ones from the original sim download) and restart the sim .exe.
* To properly sim Desert Fairy, you should activate both components in the strategy skill section in this order: "Desert 1 (negation)" and "Desert 2 (debuff)". The  latter applies the debuff to you and the enemy, the former negates the effect on your echelon.
* The Fatigue debuff should be set to skill level 10, and the GospelAxis debuff should have its skill level field left blank.

# To Do:
* Possible code/formatting clean up

# Other Links:
* KR Battle Sim: https://gall.dcinside.com/mgallery/board/view?id=micateam&no=1506585
* GamePress guide: https://gamepress.gg/girlsfrontline/how-use-gf-battle-tester-girls-frontline-battle-tester
* EN Battle Sim: https://github.com/neko-gg/gfl-combat-simulator
* Data Source: https://github.com/randomqwerty/GFLData

# Contact Info:
* Email: randomabc123456@gmail.com
* Discord: randomqwerty
* Reddit: /u/UnironicWeeaboo
