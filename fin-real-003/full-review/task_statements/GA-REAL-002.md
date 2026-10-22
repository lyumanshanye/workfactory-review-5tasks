# 审核Wesnoth忠诚派升级路线，找出表面数值比较会遗漏的代价

你为《The Battle for Wesnoth》的数值策划整理1.18.0版Loyalists阵营。团队需要一份真实配置依赖完整的招募/升级审查，决定哪些分支值得做后续游戏内测试。不能用编造敌军、权重或战斗胜率代替官方配置，也不要求在本题改游戏平衡。

## 真实来源

`/app/input_files/` 保留官方Git仓库1.18.0 tag下的原始路径，含阵营、单位、移动类型、宏和相关实现。文件内容未经修改。图片/音频不在分析范围；缺少媒体不影响静态配置审查。官方 `data/tools/wesnoth` 中的WML读取工具也作为原件提供，可用于解析。不得直接执行来源中不相关的脚本。

## 工作范围

1. 从 `data/multiplayer/factions/loyalists-default.cfg` 确定真实recruit与leader集合，区分可招募、可选领袖和升级可达单位。递归追踪全部recruit的advances_to，直到无常规升级，报告多分支、交汇、未解析ID及循环；不得把领袖列表直接当招募列表，或只读取human-loyalists目录遗漏法师/骑士/人鱼。
2. 为每个可达单位给出来源路径、ID、等级、种族、cost、HP、movement、movement_type、alignment、experience和直接武器。正确识别[unit_type]直接属性及真正[attack]，不把动画/filter_attack、文本引用或variation的字段重复当作基础武器。继承/宏不能当普通数值或静默填默认值；对确需展开的地方给出可核对的依赖说明。
3. 给每个可达单位列出flat、forest、village、shallow_water、deep_water的移动成本和防御配置，以及blade/pierce/impact/fire/cold/arcane抗性。沿实际movement_type关联公共默认，再应用单位自身覆盖；保留原始值与来源。解释防御值的方向、负值上限和抗性与承受伤害比例的区别；不要把缺失配置擅自写成0。确实需引擎规则才能解释的字段，引用随包实现，不伪称已运行引擎。
4. 对每条升级边计算HP、movement、cost和各直接武器的变化，按名称+攻击距离+伤害类型对齐武器，新增/删除武器单列。列出全部“HP上升但movement下降”的边及全部alignment改变的边。对每个招募起点到每个可达终点枚举不同简单升级路径，报告途中每个非终点的experience之和及最小值、对应路径；该和只是基础配置XP门槛，不包含击杀获益、经验倍率、traits或AMLA，不能称作真实升级耗时。
5. 审查所有可达单位的武器特殊能力和单位能力，沿官方宏定位定义。选择至少三条实际存在、会使“面板damage×number÷cost”排序失真的配置依赖，举真实单位与升级边说明。区分条件生效的charge、firststrike、magical等规则，不把特性名称当无条件数值奖励，不用静态数据推断完整对战胜率。
6. 做完整性核对：所有招募ID、升级边目标、movement_type、业务相关能力宏均有对应来源或明确未解析原因。备忘录推荐三个具体的游戏内验证问题，每个绑定实际分支、现有数值证据和需要实测的条件，不虚构现成测试日志。

## 交付

仅在 `/app/output/` 交付 `audit.xlsx`（单位目录、升级边、路径XP、地形抗性、能力依赖与覆盖）、中文 `memo.docx`、`results.json`、`analysis.py`。所有表与JSON保留源路径/标签/行号；排名/路径运算使用原始值。`python analysis.py --input /app/input_files --output /app/output` 离线重建另外三份文件。允许标准库、openpyxl、python-docx、lxml及随包官方WML读取工具；不能硬编码结果、修改输入或以原始配置不存在的游戏规则凑答案。
