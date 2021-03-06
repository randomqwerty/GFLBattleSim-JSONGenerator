# GFLBattleSim-JSONGenerator

This is a macro-enabled Excel spreadsheet that will generate/modify the following JSON files needed for the GFL KR Battle Sim:

* gun_with_user_info.json
* fairy_with_user_info.json
* equip_with_user_info.json
* GFBattleSimulator.json (modified, file must already exist)
* mission_act_info.json (modified, file must already exist)

Specifically, this can be used to:

* Change battle settings (day/night, enemy ID, boss HP, enable/disable HOC/Fairy)
* Edit echelon composition (dolls, skill levels, equipment, positions, etc.)
* Edit fairy information (fairy, rarity, level, skill level, etc.)
* Apply strategy fairy skills (Parachute, Construction, Suee, Combo)
* Apply enemy debuffs (currently only EMP)

This does NOT adjust the HOC JSON file (squad_with_user_info.json).

Battle Sim can be downloaded from here: https://gall.dcinside.com/mgallery/board/view?id=micateam&no=1506585

VBA code is provided in a separate .vb file so it can be examined. Place the spreadsheet in the parent directory of the battle sim (i.e., the same folder as GFBattleSimulator.json).
