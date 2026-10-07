# 心房搜查令

第十四届校园心理剧大赛参赛作品，2024 级数据科学与大数据技术班。

剧本借了《人民的名义》里搜查那段做框架，改成一场"心房搜查"。室友兼心理委员郝朋打着搜查的旗号，一点点撬开赵学衡藏着不肯给人看的东西，从强装镇定到防线崩掉，最后落到接纳自己。全剧八幕。

## 东西都在哪

- `剧本/`　Word 源稿、交上去的 PDF 版，还有一份从文档里抽出来的纯文字台词
- `参赛材料/`　比赛通知，还有三个附件（规则、报名表、剧本格式要求）
- `素材/`　写剧本时参考的字幕文件
- `脚本/`　生成放映 PPT 用的几个 Python 脚本
- `剧照/`　八张分镜剧照
- `预览图/`　放映版导出的十一张图片，不想开 PPT 的时候直接翻这个

根目录下面两个 pptx：`心房搜查令_最终现场放映版.pptx` 是上台放映用的正式版；带"电影分镜剧照与台词字幕"字样的那个是中间模板，用脚本可以重新生成。

## 放映 PPT 是怎么来的

PPT 不是一张张手做的。剧照和台词是脚本拼进去的，改台词改剧照都只要重跑一遍，比手动调版式省事。

先在项目根目录开个终端，然后按顺序跑：

```bash
cd D:\develop\心理

python 脚本\build_template_ppt.py     # 先生成中间模板
python 脚本\run_agent_pipeline.py     # 把 剧照\ 里八张图注入模板，输出正式放映版
python 脚本\verify_final_ppt.py       # 检查页数、画幅、图层顺序、场控表有没有出问题
python 脚本\export_all_slides.py      # 导出 预览图\（得装了 PowerPoint 才行）
```

其中 `export_all_slides.py` 和 `export_png.py` 是调 PowerPoint 的 COM 接口导图的，只能在装了 Office 的 Windows 上跑。`test_preview.py` 是早期试封面效果留下的草稿，平时用不上。

## 几个坑

- 脚本里的路径全是相对"当前目录"写的，所以**一定要先 cd 到项目根目录**再执行，不然会找不到文件。
- `剧照/` 这个目录名在脚本里是写死的，改名脚本就找不到图了。
- 预览图和 `__pycache__` 这些随时能重新生成，所以没放进 git。

## git

Gitee 和 GitHub 各存了一份，内容一样。改完东西：

```bash
git add -A
git commit -m "这次改了什么"
git pushall
```

`pushall` 是本机配的别名，等于依次推 gitee 和 github；换台机器就没这个别名了，得写成 `git push gitee master` 和 `git push github master` 两条。
