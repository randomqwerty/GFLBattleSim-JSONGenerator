# GFLBattleSim-JSONGenerator

This is a macro-enabled Excel spreadsheet that will generate/modify the following JSON files needed for the GFL KR Battle Sim:

* gun_with_user_info.json
* fairy_with_user_info.json
* equip_with_user_info.json
* squad_with_user_info.json
* chip_with_user_info.json (this will be a blank file)
* GFBattleSimulator.json (this is modified, the file must already exist)
* mission_act_info.json (this is modified, the file must already exist)

Specifically, this can be used to:

* Change battle settings (day/night, enemy ID, boss HP, enable/disable HOC/Fairy)
* Edit echelon composition (dolls, skill levels, equipment, positions, etc.)
* Edit fairy information (fairy, rarity, level, skill level, etc.)
* Edit HOC information (stats, skill level, whether to use said HOC, etc.)
* Apply strategy fairy skills (Parachute, Construction, Suee, Combo)
* Apply enemy debuffs (currently only EMP)

Note that you should enter the FINAL stats if simming HOCs (i.e., including chips and iterations). The base stats will be adjusted to get the desired stats, the chip boards will be blank and they will appear as max iteration in-game.

Battle Sim can be downloaded from here: https://gall.dcinside.com/mgallery/board/view?id=micateam&no=1506585

VBA code is provided in a separate .vb file so it can be examined. Place the spreadsheet in the parent directory of the battle sim (i.e., the same folder as GFBattleSimulator.json).
