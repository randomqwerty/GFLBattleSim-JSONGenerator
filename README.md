# GFLBattleSim-JSONGenerator
<img src="https://i.imgur.com/qIBZMB6.png" width="100%">

**IMPORTANT NOTE: Due to the introduction of client v2.09 and the equipment index, you will need to replace the default userinfo.json with the one in this repo. This version adds the equip_collect field that the client is expecting.**

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
* As mentioned above, please replace the default userinfo.json with the one in this repo. This new version adds the equip_collect field (used for the equipment index) and also enables Unity Skills by adding items IDs 1100001-1100006 to item_with_user_info.
* When entering G&K HOC stats, use the FINAL stats including chips and iteration bonuses. The base stats will be adjusted to get the desired stats, so your HOCs will appear to be max iteration and have blank chip boards.
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
