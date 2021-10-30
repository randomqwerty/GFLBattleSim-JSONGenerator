# GFLBattleSim-JSONGenerator
<img src="https://i.imgur.com/qIBZMB6.png" width="100%">

This is a macro-enabled Excel spreadsheet that will generate/modify the JSON files needed for the GFL KR Battle Sim.

Specifically, this can be used to:

* Change battle settings (day/night, enemy ID, boss HP, enable/disable HOC/Fairy)
* Edit G&K echelon composition (dolls, skill levels, equipment, positions, etc.)
* Edit SF echelon composition (ringleader/mooks, skill levels, positions, etc.)*
* Edit fairy information (fairy, rarity, level, skill level, etc.)
* Edit HOC information (stats, skill level, whether to use said HOC, etc.)**
* Apply strategy fairy skills (Parachute, Construction, Suee, Combo, etc.)
* Apply enemy debuffs (currently only EMP, Auspicious)

To enable unity skills, replace the default userinfo.json with the one in this repo. This new version simply has item IDs 1100001 to 1100006 added to the "item_with_user_info" section.

\** When simming HOCs, you should enter the FINAL stats if simming HOCs (i.e., including chip and iteration stats). The base stats will be adjusted to get the desired stats, so your HOCs will appear to be max iteration with blank chip boards.

VBA code is provided in a separate .vb file so it can be examined. Place the spreadsheet in the parent directory of the battle sim (i.e., the same folder as GFBattleSimulator.json).

# To Do:

* Add other strategy fairy skills (e.g., Desert)
* Make this less janky (maybe)

# Other Links:
* KR Battle Sim: https://gall.dcinside.com/mgallery/board/view?id=micateam&no=1506585
* GamePress guide: https://gamepress.gg/girlsfrontline/how-use-gf-battle-tester-girls-frontline-battle-tester
* EN Battle Sim: https://github.com/neko-gg/gfl-combat-simulator

# Contact Info:
* Discord: Randomqwerty#4678
