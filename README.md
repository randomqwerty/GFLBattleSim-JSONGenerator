# GFLBattleSim-JSONGenerator
<img src="https://i.imgur.com/qIBZMB6.png" width="100%">

This is a macro-enabled Excel spreadsheet that will generate/modify the JSON files needed for the GFL KR Battle Sim.

Specifically, this can be used to:

* Change battle settings (day/night/target practice, enemy ID, boss HP, HOC/Fairy, etc.)
* Edit G&K echelon composition (dolls, skill levels, equipment, positions, etc.)
* Edit SF echelon composition (units, skill levels, chips, positions, etc.)
* Edit fairy information (fairy, rarity, level, skill level, etc.)
* Edit HOC and SF HOC information (stats, skill level, etc.)
* Apply strategy fairy skills (Parachute, Construction, Suee, Combo, etc.)
* Apply enemy debuffs (currently only EMP, Auspicious)
* Load/save/delete preset echelon

VBA code is provided in a separate .vb file so it can be examined. Place the spreadsheet in the parent directory of the battle sim (i.e., the same folder as GFBattleSimulator.json).

# Additional Usage Notes:
* Unity Skills can be enabled by replacing the default userinfo.json with the one in this repo. This new version simply has item IDs 1100001 to 1100006 added to the "item_with_user_info" section.
* When entering G&K HOC stats, use the FINAL stats including chips and iteration bonuses. The base stats will be adjusted to get the desired stats, so your HOCs will appear to be max iteration and have blank chip boards.
* You cannot sim more than 8 G&K HOCs at a time. This appears to be an issue with the sim itself.
* There are issues with simming G&K directly after simming SF echelons. If you want to switch back to G&K, you should generate the new JSONs (or use the ones from the original sim download) and restart the sim .exe.

# To Do:
* Add other strategy fairy skills (e.g., Desert)
* Make this less janky (maybe)

# Other Links:
* KR Battle Sim: https://gall.dcinside.com/mgallery/board/view?id=micateam&no=1506585
* GamePress guide: https://gamepress.gg/girlsfrontline/how-use-gf-battle-tester-girls-frontline-battle-tester
* EN Battle Sim: https://github.com/neko-gg/gfl-combat-simulator
* Data Source: https://github.com/randomqwerty/GFLData

# Contact Info:
* Discord: Randomqwerty#4678
