# GFLBattleSim-JSONGenerator
![screenshot](https://i.imgur.com/qIBZMB6.png)

This is a macro-enabled Excel spreadsheet that will generate/modify the following JSON files needed for the GFL KR Battle Sim.

Specifically, this can be used to:

* Change battle settings (day/night, enemy ID, boss HP, enable/disable HOC/Fairy)
* Edit G&K echelon composition (dolls, skill levels, equipment, positions, etc.)
* Edit SF echelon composition (ringleader/mooks, skill levels, positions, etc.)
* Edit fairy information (fairy, rarity, level, skill level, etc.)
* Edit HOC information (stats, skill level, whether to use said HOC, etc.)
* Apply strategy fairy skills (Parachute, Construction, Suee, Combo)
* Apply enemy debuffs (currently only EMP)

Note that you should enter the FINAL stats if simming HOCs (i.e., including chips and iterations). The base stats will be adjusted to get the desired stats, the chip boards will be blank and they will appear as max iteration in-game.

Battle Sim can be downloaded from here: https://gall.dcinside.com/mgallery/board/view?id=micateam&no=1506585

VBA code is provided in a separate .vb file so it can be examined. Place the spreadsheet in the parent directory of the battle sim (i.e., the same folder as GFBattleSimulator.json).

Contact me at Randomqwerty#4678 on Discord to report bugs or provide feedback.
