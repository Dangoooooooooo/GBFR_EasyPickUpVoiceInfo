import tkinter as tk
import threading
from tkinter.scrolledtext import ScrolledText
import sys
from A2M1 import Action2Motion_Process
from M2E1 import Motion2Event_Process
from E2wID1 import Event2wemID_Process
from AltWavName1 import AltWavName_Process
from bnkRepack1 import bnkRepack_Process
from MixUp11 import MixUp_Process
from tkinter import scrolledtext

title_text_Eng = r"Granblue Fantasy:Relink - Easy PickUp Characters' VoiceInfo <BY:Dangoooooo | BiliBili  QQ:1041271418> V1.0.2"

button_texts_Eng = ['(#1) PickUp [Action] <> [Motion]',
                '(#2) PickUp [Motion] <> [EventName]',
                '(#3) PickUp [EventName] <> [.wem]IDs',
                '(#4) Mix Them All',
                '(#5) Alter [.wav]s\' name to [EventName]s',
                '(#6) Repack [.bnk]s with new [.wem]s']

First_text_Eng = r'''First, use "GBFRDataTools" to unpack the "data.i" in the game directory <.\Steam\steamapps\common\Granblue Fantasy Relink>.
* You will get a "data" folder in the same directory.
* Read the "Preparation Required" instructions on the right.
* After the preparation work is completed, press the button on the left.
>> Download "GBFRDataTools" at <github.com/Nenkai/GBFRDataTools>.
>> This tool originates from <github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>. Feel free to ask questions or provide feedback.'''

indicates_Eng = [
r'''%% Preparation Required %%
1) Copy <data\system\player\data\plxxxx\plxxxx_action.msg> to <.\Yours\> folder.''',
r'''%% Preparation Required %%
1) Copy <data\pl\plxx00\> folder to <.\Yours\> folder.''',
r'''%% Preparation Required %%
1) Complete #2.
2) Copy <data\sound\(Japanese or English(US))\vo_plxx00(_m, _0x_0x, _town, etc.).(bnk or pck)> to <.\Yours\> folder.
* "vo_plxx00.bnk" is a HIRC file. Do not put it in <.\Yours\> folder, otherwise it will cause an error.''',
r'''%% Preparation Required %%
1) Complete #1, #2, #3.
2) Recommended to install "GBFR Logs" and copy <..\GBFR Logs\lang\(your preferred language)\ui.json> to <.\Yours\> folder.
* Recommended to only process [.bnk] audio files.
* If necessary, use "RingingBloom" to unpack [.pck] audio files to <.\Yours\wem_org\(a certain pck file full name)\> folder.
>> Download "GBFR Logs" (DPS plugin) at <github.com/false-spring/gbfr-logs>.
>> Download "RingingBloom" at <github.com/Silvris/RingingBloom>.''',
r'''%% Preparation Required %%
1) Fill in the "YourRecord" table in "GBFR#0{character number}{character name}VoiceInfo_MixUp.xlsx".
2) Create a "wav_org" folder in <.\Yours\> folder. Put your favorite [.wav] files into it.
* You must open, edit, save, and close!!!''',
r'''%% Preparation Required %%
1) Complete #5.
2) Put the [.wem] files you converted from "Wwise Launcher" into <.\Yours\wem_FromWwise(a certain bnk file full name) folder>
* The specific conversion operation steps can be seen in Q7 of "Q&A".
* After completion, you can get your repackaged [.bnk] file in <.\Yours\bnkEXPORT\>.
>> Download "Wwise Launcher" at <audiokinetic.com>.''']

qnatext_Eng = r'''This tool is designed for creating voice mods for the game "Granblue Fantasy: Relink". It conveniently extracts the relationships between each character's "Action", "Motion", "EventName", and the corresponding "wemID" in the bnk files, and generates an Excel table. Additionally, after filling out the table, it can repackage the content into the corresponding bnk files. If you have any questions, you can contact me on BiliBili by searching for Dangoooooo, or find me through QQ number: 1041271418. You can also find the source code and the newly released tool package directly on GitHub at <http://github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>. I will check periodically.

Please read the following Q&A section carefully, as well as the comments on the GBFR#EasyPickUpVoiceInfo.exe tool.

If you have any questions, you can search for Dangoooooo on the BiliBili website to contact me. You can find me through the qq number: 1041271418, or you can directly find the source code and newly released tool package at <http://github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>. I will check from time to time.

>>>Q&A<<<

Q1: Why find the corresponding relationship?
A1: Because the character voices in this game may be reused, the multiple moves you make by issuing commands may correspond to the same group of voices, which will bring us a lot of trouble in making voice mods.

Q2: What is the specific corresponding relationship?
A2: The corresponding relationship is as follows:
1. Action refers to the move (we issue commands to control the character to make the move) corresponding to multiple Motions (action combinations) [recorded in the plxxxx.msg file]
2. Each Motion corresponds to multiple EventNames (including sound, mouth movement, facial expression, etc.) [recorded in the bxm series files in the pl\plxxxx folder, we only care about _se.bxm files, which records the EventName related to sound]
3. Each wem file has a corresponding ID, and its serial number (NumSerial) in the bnk file, and the EventName related to sound will be engraved in the wem file (you need to open it in hexadecimal to view and extract it separately, such as 010Editor). It should be noted that the sound EventName recorded in the bxm file and the one engraved in the wem are slightly different. The specific performance is that the EventName in the wem will become like 01, 02 after the 4-digit character number after pl according to the game update, and there will be _1, _2 or _a, _b suffixes at the end, while the EventName recorded in the bxm is more like a summary information. In summary, the corresponding form is wemID---NumSerial---EventName(with Serial) one-to-one correspondence, and multiple EventName(with Serial) will correspond to one EventName(in .bxm)
4. Multiple wems will be packaged into multiple bnks or pcks according to the content and updates of the game. [In the bnk rules of GBFR, the vo_plxxxx.bnk file is used as the total control HIRC file, without sound content. All others have sound content, all are added with suffixes similar to _01_01 or _m or _town after the HIRC file, or their mix]

Q3: What needs to be prepared in advance?
A3: The preparations are as follows:
1. First, you need to use "GBFRDataTools" <http://github.com/Nenkai/GBFRDataTools> to unpack the "data.i" file in the game root directory <.\Steam\steamapps\common\Granblue Fantasy Relink>. You will get a huge "data" folder in the same directory (it is recommended to copy and rename the previous "data" folder to "data_org"), which contains all the resources of the entire game, its size is almost the same as the size of the game body, at least 70G of hard disk space should be reserved.
2. Your computer needs to have "Wwise Launcher" installed (download wwise after downloading from <http://www.audiokinetic.com>), it is used to help you convert wav to wem files.

Q4: What do I need to fill in after generating the table?
A4: We only care about the "VoiceInfo_MixUp" table, what needs to be filled in is the ABCD column in the "YourRecord" sheet in the table, where the second row has instructions and the third row has demonstrations. Specifically, the B column fills in the serial number you want the original audio to correspond to ("belonging to bnk|pck package"_"digital number"), and the C column fills in the wav file name of the new audio you want to replace (do not fill in the file extension too much). And the A column is the name you give to this wem corresponding to the action, you can delete it yourself, or you can extract the move name from the blood display plugin GBFR logs, or you can ignore it directly. The D column is a note that is completely filled in according to your own needs. The remaining columns are automatically filled in with built-in formulas.

Q5: So how do I find the serial number corresponding to the original audio?
A5: The belonging bnk|pck package can be found in the explanation in the second row of the B column, and there are two ways to know the following digital number:
1. You can directly check the "MixUp" sheet in the "VoiceInfo_MixUp" table, find out which wem audios they correspond to through the known Action number or name, or if there is no corresponding Action and Motion information, you can also infer the function of this audio through the specific content of the "EventName" under the "PL_vo_Event_WithSerial" column.
2. You can also generate AI voices on the website (there should be many, I recommend a free website in China <https://ttsmaker.cn/>), let it read the number (such as reading a245, b431, etc.), and then use the slicing tool <http://github.com/flutydeer/audio-slice> to cut it open, and use this tool software to pack it back into the game. You will be able to hear very intuitively which wem file your move corresponds to. But the disadvantage is that many times a certain action will correspond to multiple wem files, maybe you need to repeat this action many times to listen to the whole. I also provide the Chinese and English versions of the sliced number compression package "Serial(unpack to wav_org folder)" in the release of <http://github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>, from a to j, each from 0 Count to 2000 (because English letters may be confused, I will add some changes to it, it should be enough).

Q6: Why not unpack and pack pck with one click?
A6: Because the pck audio file involves things like audio looping, you need to set the loop point yourself, and most of the battle voices are in the bnk file (except for the vo_plxxxx.bnk which is the total control HIRC), the most important thing is that I did not find a tool on the Internet that can replace the wem file in pck in batches, and I will do it when I have the opportunity. But when I generate the comparison table, I also consider the pck file. You can use the "RingingBloom" <http://github.com/Silvris/RingingBloom> software to unpack and pack the pck file.

Q7: I hope that after the replacement, the character's lips will open when reciting lines. What should I do?
A7: This problem is essentially to engrave the correct EventName in the new wem. The specific operation is:
1. Open Wwise Launch to create a new project, then select "Share Sets">"Conversion Settings">"Default Conversion Settings" in the upper middle of the left column of Wwise Launcher and double-click to open it, find "Insert filename marker" on the right, and check it.
2. Select "Audio">"Interactive Music Hierarchy" in the upper middle of the left column of Wwise Launcher, right-click in the "Default work unit" below, select "Import Audio Files..." to add the "wav_AltName" audio folder generated by this tool software (#5) in the Yours folder, and directly right-click here to select "Convert..." to convert all the renamed wavs into wem with EventName. The converted wem file is in <This PC\Document\WwiseProjects\(your project name)\.cache\Windows\SFX\>"wav_AltName".
3. Rename the "wav_AltName" in the SFX output folder to "wem_FromWwise", and then throw the entire folder back to <.\Yours\>, and then run (#6) step.

Q8: I hope to modify the corresponding relationship between Motion and EventName in the bxm file. Can it be done?
A8: Of course it can. The operation steps are as follows:
1. By completing the (#2) step of this tool software, you can not only get the "Motion_EventName_ComparisonTable" comparison table in the "Yours" folder, but also get the series of files that have been unpacked from bxm to xml in <.\Yours\plxxxx\_se_bxm_xml_Group\>.
2. You can easily open the xml file with a notepad, and find that it is composed of many "Seq LayerFlag" entries. I intercepted one of the entries, something like this: <Seq LayerFlag="4294967295" StartTime="0.016667" SeqFlag="0" EventName="PL1100_vo_ATK_default_s" PartsNo="0" Flag="0" OffsetX="0.000000" OffsetY="0.000000" OffsetZ="0.000000"/>.
3. We can perform the "delete" operation on the entire line of entries in it, but we need to pay attention to modify the "SeqNum" at the top of the xml file to be equal to the current total number of entries.
4. We can also "replace" the "EventName" parameter in the entry with other existing "EventName" in the bxm file. It is very important to note that if you use the "replace" operation, then the duration of the wem group corresponding to the new "EventName" must be shorter than any of the old wem group! (Generally, most of our wem groups are replaced with the same voice) Otherwise, when a certain action A triggers this voice, it will make the voice of the next derivative action B unable to play! If the voice you need is very long and hard to give up, there are two solutions: (a) modify the "StartTime" parameter in the entry, reduce it to 0, can extend the playback time to a certain program; (b) is to modify the "PartsNo" parameter in the entry, if its value and another bxm file after the derivative action The value of "PartsNo" is different (the general value is "0" and "1"), which means that they belong to different "channels", that is, the long voice triggered by a certain action A will play together with the voice of the next derivative action B! It can also solve the problem of long voice to a certain program.
5. Finally, the modified xml file can be packed into a bxm file using the opening method of <.\Yours\nier_cli_mgrr\nier_cli_mgrr.exe>, or directly drag the xml file into the exe.

Q9: Will there be some actions that use individual audios in other Event groups? If so, what should I do?
A9: This situation really exists. This may occur in the audio corresponding to the link attack. When this situation occurs, we need to modify the total control HIRC file, that is, the vo_plxxxx.bnk file, which records the call rules. How to modify it can be explained specifically through an example: When I was doing Siegfried voice mod before, I found that the digital serial number "408","786","975","1142" in the "PL1100_vo_ATK_default_l" Event group Among them, "786" and "1142" were requisitioned to go to the link attack (this can only be known by listening to the number repeatedly). So I need to use the 010Editor hexadecimal editor to open the HIRC file. After opening, use the unsigned integer (ui32) to search and use the "label" to record all the positions of the wemID of the entire Event group in this HIRC file, that is, the digital serial number "408","786","975","1142" corresponds to the wemID "308361897","571928584","696541074","798788669" in this HIRC file All positions. We will find that the wemID of the 4 members of this Event group appears compactly in 3 address blocks of the HIRC file, and only the wemID corresponding to "786" and "1142" also appears compactly in two other places. Then we only need to replace the wemID corresponding to these two groups of "786" and "1142" that appear alone with the wemID of other members that originally only exist in the link attack group, save and get the new HIRC file.

Q10: How to pack the new bnk into a mod and use it after conversion?
A10: The specific steps are as follows:
1. Go to this website <http://nenkai.github.io/relink-modding/getting_started/mod_manager/> to install "Reloaded-II Mod Manager" and "gbfrelink.utility.manager" and make sure they are the latest versions.
2. Run the Reloaded-II mod manager, find the "3 gear icon" on the upper left, and then click the "Add" button on the right, then fill in the mod information according to the actual needs. After filling in, you will find the newly created mod folder in the "Mods" in the Release folder of the mod manager, and then create a series of folders according to the specific location of your bnk|pck file relative to the "data" unpacked file, and then throw the bnk|pck file into it. For example, my new bnk file is placed in <.\Release\Mods\gbfr.voice.Siegfried_DARKv1.1.2\gbfr.voice.Siegfried_DARK\GBFR\data\sound\Japanese\>.
3. Find the "Add Application" on the upper left of Reloaded-II, select the exe file in the game root directory, and then the mod will appear on the right, click the square plus sign in front to use it.
'''

title_text_Ch = r"碧蓝幻想:Relink - 简易提取角色语音【BY:Dangoooooo | BiliBili  QQ:1041271418】V1.0.2"

button_texts_Ch = ['(#1) 提取[Action]与[Motion]关系',
                '(#2) 提取[Motion]与[EventName]关系',
                '(#3) 提取[EventName]与[.wem]ID关系',
                '(#4) 提取结果全部混合',
                '(#5) 将[.wav]文件全改名为[EventName]',
                '(#6) 将新[.wem]文件组重新打包成[.bnk]']

First_text_Ch = r'''首先，使用"GBFRDataTools"解包游戏根目录<.\Steam\steamapps\common\Granblue Fantasy Relink>中的"data.i"。
* 你将在同一目录下得到 "data" 文件夹。
* 阅读右侧的 "需准备" 指令。
* 准备工作完成后，按左侧的按钮。
>> 在<github.com/Nenkai/GBFRDataTools>下载 "GBFRDataTools"。
>> 该工具源自 <github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>。欢迎提问或反馈。'''

indicates_Ch = [
r'''%% 需准备 %%
1) 复制<data\system\player\data\plxxxx\plxxxx_action.msg>到<.\Yours\>文件夹。''',
r'''%% 需准备 %%
1) 复制<data\pl\plxx00\>文件夹到<.\Yours\>文件夹。''',
r'''%% 需准备 %%
1) 完成#2。
2) 复制<data\sound\(Japanese或English(US))\vo_plxx00(_m、_0x_0x、_town等等).(bnk或pck)>到<.\Yours\>文件夹。
* "vo_plxx00.bnk"是HIRC文件。不要把它放在<.\Yours\>文件夹里，否则会报错。''',
r'''%% 需准备 %%
1) 完成#1,#2,#3。
2) 建议安装"GBFR Logs"并复制<..\GBFR Logs\lang\(你喜欢的语言)\ui.json>到<.\Yours\>文件夹。
* 建议只处理[.bnk]音频文件。
* 如果很必要，使用"RingingBloom"解包[.pck]音频文件到<.\Yours\wem_org\(某个pck文件全名)\>文件夹。
>> 在<github.com/false-spring/gbfr-logs>下载"GBFR Logs"(DPS插件)。
>> 在<github.com/Silvris/RingingBloom>下载"RingingBloom"。''',
r'''%% 需准备 %%
1) 在"GBFR#0{角色编号}{角色名}VoiceInfo_MixUp.xlsx"中填写"YourRecord"表格。
2) 在<.\Yours\>文件夹中创建"wav_org"文件夹。将你自己所喜欢的[.wav]文件放进去。
* 必须打开、编辑、保存、关闭!!!''',
r'''%% 需准备 %%
1) 完成#5。
2) 将你从"Wwise Launcher"转换的[.wem]文件放到<.\Yours\wem_FromWwise(某个bnk文件全名)文件夹>
* 具体转化操作步骤可以看"Q&A"里的Q7。
* 完成后你可以在<.\Yours\bnkEXPORT\>中得到你重新打包的[.bnk]文件。
>> 在<audiokinetic.com>下载"Wwise Launcher"。''']

qnatext_Ch = r'''这个工具是给 《Granblue Fantasy: Relink》 这个游戏做语音mod用的，它能很方便地提取出游戏里每个角色的“Action”、“Motion”、“EventName”、以及对应bnk文件里“wemID”的对应关系，并生成Excel表。此外也能在填写完表格后，根据其内容重新打包成对应bnk文件。

>>>工作流程如下<<<
(#1) 从plxxxx_action.msg文件中提取出“Action”和“Motion”的对应关系，并生成表。
(#2) 在pl\plxxxx文件夹的后缀为“_se.bxm”文件组中提取出“Motion”与”EventName“的对应关系，并生成表。
(#3) 在vo_plxxxx(_m,_0x_0x,_town,...).(bnk,pck)文件中（可以多个）提取出“EventName”与“wemID”的关系，同时依照“wemID”序号大小产生编号，并生成表。
(#4) 将3个表融合在一起为总表。
(#5) 将你自己的wav音频放到<.\Yours\wav_org>文件夹里（没有就新建）。工具会根据总表“YourRecord”sheet内C列信息，从文件夹里找到相同wav音频，并将其名字替换为H列的“EventName”（目的是为了能让嘴唇动起来）,最终输出在”wav_AltName“文件夹里。
(#6) 在Wwise Launcher工具里导入整个"wav_AltName"文件夹，全部转为wem格式，在<文档\WwiseProjects\（你的项目名）\.cache\Windows\SFX>输出文件夹中把"wav_AltName"文件夹改名为“wem_FromWwise”，并扔回“Yours”文件夹里。工具会将新wem文件组重新封包bnk文件到<.\Yours\bnkEXPORT\>文件夹里。

请仔细阅读下面的Q&A版块，以及GBFR#EasyPickUpVoiceInfo.exe工具上的注释。

如果有任何问题，都可以在BiliBili网站搜Dangoooooo联系我。可以通过qq号：1041271418找到我，也能直接在github上找到源代码和新发布的工具包 <http://github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>，我会不定时查看。

>>>Q&A<<<

Q1: 为什么要找出对应关系？
A1: 因为这个游戏的角色语音会存在复用的情况，你发出指令搓出的多个招有可能会对应同一组语音，会给我们做语音mod带来不少困扰。

Q2: 对应关系具体是怎么样的呢？
A2: 对应关系如下：
1. Action就是指招数（我们发出指令控制人物发出的招数）对应多个Motion（动作组合）[ 记录在plxxxx.msg文件里 ]
2. 每个Motion又对应多个EventName（其中有声音、嘴巴动作、脸部表情等等）[ 记录在pl\plxxxx文件夹里的bxm系列文件，我们只关心_se.bxm文件，它里面记录着与声音相关的EventName ]
3. 每个wem文件都有对应的ID，以及它在bnk文件里的序列号（NumSerial），且和声音相关的EventName会刻在wem里文件内部（要用16进制方式打开方可单独查看和提取，比如010Editor）。要注意的是记录在bxm文件里的声音EventName和刻在wem里的略微不一样。具体表现是wem内的EventName在pl后的4位角色编号后2位会根据游戏更新变成像01、02这样，同时在末尾会出现_1、_2或者_a、_b这样的后缀，而bxm内记载的EventName更像一个汇总信息。综上所述，对应形式是wemID---NumSerial---EventName(with Serial)一一对应，而多个EventName(with Serial)会对应一个EventName(in .bxm)
4. 多个wem会根据游戏的内容和更新封装成多个bnk或者pck。[ GBFR的bnk制定规则里vo_plxxxx.bnk文件是用作总控的HIRC文件，没有语音内容。除此以外的都是具有语音内容的，都是在HIRC文件后面加类似于_01_01或者_m或者_town的后缀，亦或是它们的混合]

Q3: 有什么需要事先准备的？
A3: 准备事项为：
1. 首先你需要用"GBFRDataTools"<http://github.com/Nenkai/GBFRDataTools>解包在游戏根目录下<.\Steam\steamapps\common\Granblue Fantasy Relink>的 "data.i" 文件。你会在同目录下得到一个超大的"data"文件夹（推荐把之前的“data”文件夹复制改名备份成"data_org"）里面包含了整个游戏所有的资源，它的大小和游戏本体大小差不多，至少要预留70个G的硬盘空间。
2. 你的电脑里需要装有“Wwise Launcher”(从<http://www.audiokinetic.com>下载后安装wwise)，它是用来帮你把wav转化为wem文件。

Q4: 生成表后我具体要填写什么呢？
A4: 我们只关心"VoiceInfo_MixUp"表，需要填写的是表里的"YourRecord"sheet内的ABCD列，其中第2行有说明，第3行有演示。具体是B列填写你想原音频所对应序号（"所属bnk|pck包"_"数字编号"），C列填写你想替换成的新音频的wav文件名（不要多填入文件扩展名）。而A列是你给这个wem名字对应的动作起的名字，你既可以删掉自己填，也能根据我从显血插件GBFR logs里提取出的招数名，还能直接忽视不填。D列是完全根据自己需求填写的备注。剩余列则内置公式自动填充。

Q5: 那我该怎么找到原音频所对应的序号呢？
A5: 所属bnk|pck包可在B列第2行的说明里找到字母代号，而后面的数字编号有两种办法得知：
1. 可以直接查阅"VoiceInfo_MixUp"表里的“MixUp”sheet内，通过已知Action编号或者名字来查找它们对应哪些wem音频，或者如果没有对应的Action和Motion等信息，也能通过“PL_vo_Event_WithSerial”列下的“EventName”的具体内容来推测该音频的作用。
2. 你也可以在网站（应该挺多的，我这边推荐一个中国内免费的网站<https://ttsmaker.cn/>）上生AI语音，让它读编号（比如读a245，b431等），再利用切片工具<http://github.com/flutydeer/audio-slice>切开，并利用这个工具软件封包回游戏内。你就能很直观的听到你的招数会对应哪个wem文件了。不过缺点是很多时候某个动作会对应多个wem文件，也许你需要重复多次这个动作才能听完整。我这边也在<http://github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>的release里提供了中文版和英文版已经切片好的编号压缩包“Serial(unpack to wav_org folder)”，从a到j，每个从0数到2000（由于英文字母可能会听混，我会加一些变化上去，应该是够用了）。

Q6: 为什么不做pck的一键解包和封包？
A6: 因为pck语音文件内涉及到语音的循环什么的，要自己设置循环点，且绝大部分的战斗语音都在bnk文件里（除了vo_plxxxx.bnk这个总控HIRC的），最主要的是我没找到网上有能批量替换pck内wem文件的工具，有机会再弄了。不过我生成对照表的时候也考虑到了pck文件，你可以通过"RingingBloom" <http://github.com/Silvris/RingingBloom>这个软件去解包和封包pck文件。

Q7: 我希望替换完后人物的嘴唇要在念台词时张嘴怎么办？
A7: 这个问题本质上就是将正确的EventName刻在新的wem里。具体操作是：
1. 打开Wwise Launch新建一个项目，接着在Wwise Launcher里左栏靠中上部依次选择“Share Sets”>“Conversion Settings”>“Default Conversion Settings”并双击打开，在右边找到“Insert filename marker”，并勾选上即可。
2. 在Wwise Launcher里左栏靠中上部依次选择“Audio”>“Interactive Music Hierarchy”，在下面的“Default work unit”里右键选择“Import Audio Files...”添加本工具软件(#5)在Yours文件夹生成的“wav_AltName”音频文件夹，并直接在这里单击右键选择“Convert...”进行转化，即可把刚刚改好名的所有wav转为带有EventName的wem。转化完成的wem文件在<此电脑\Document\WwiseProjects\（你的项目名）\.cache\Windows\SFX\>的"wav_AltName"里。
3. 将SFX输出文件夹里的“wav_AltName”改名为“wem_FromWwise”，然后整个文件夹丢会<.\Yours\>文件里，再运行(#6)步。

Q8: 我希望能修改bxm文件里Motion与EventName的对应关系可以办得到吗？
A8: 当然可以。操作步骤如下：
1. 通过完成本工具软件的(#2)步，你就可以在“Yours”文件夹里既能得到“Motion_EventName_ComparisonTable”这个对应关系表，又能在<.\Yours\plxxxx\_se_bxm_xml_Group\>里得到已经将bxm解包为xml的系列文件。
2. 你可以轻松地通过记事本来打开xml文件，发现里面是由很多行“Seq LayerFlag”词条组成。我截取了其中一行词条，类似于这样：<Seq LayerFlag="4294967295" StartTime="0.016667" SeqFlag="0" EventName="PL1100_vo_ATK_default_s" PartsNo="0" Flag="0" OffsetX="0.000000" OffsetY="0.000000" OffsetZ="0.000000"/>。
3. 我们可以对里面整行词条进行“删除”操作，但是需要注意修改xml文件上方的“SeqNum”令其等同于当前总词条数。
4. 我们也可以对词条里的“EventName”参数“替换”为别的bxm文件里现有的“EventName”。需要非常注意的是如果你使用了“替换”的操作，那么新的“EventName”所对应的wem组音频时长一定要比旧的wem组任何一个都短！（一般我们大部分的wem组都是替换的同一个语音就是了）否则当某个动作A触发该语音后，会令下一个派生动作B的语音没法播放！如果你所需要的语音很长且难以割舍，那还有两个解决办法：(a)修改词条里的“StartTime”参数，把它调小到0，能一定程序上延长播放时间；(b)是修改词条里的“PartsNo”参数，如果它的值和派生动作后的另一个bxm文件里的“PartsNo”值不同（一般值都是"0"和"1"），那意味着它们属于不同的“频道”，也即某个动作A触发的长语音会和下一个派生动作B的语音一同播放！也能一定程序上解决长语音的问题。
5. 最后可将修改后的xml文件使用<.\Yours\nier_cli_mgrr\nier_cli_mgrr.exe>这个打开方式，或者直接将xml文件拖入exe中，即可封包成bxm文件。

Q9: 会不会有出现一些动作用了其他Event组里的个别音频？如果有又该怎么办呢？
A9: 这种情况还真有。这有可能出现在link攻击所对应的音频里。当出现这种情况时，我们是需要通过修改总控HIRC文件了，也就是vo_plxxxx.bnk文件了，它里面有记载着调用规则。具体该怎么修改，可以通过一个例子来具体说明：我之前在做Siegfried语音mod的时候发现，隶属于“PL1100_vo_ATK_default_l”这个Event组里的数字序号为"408","786","975","1142"中的"786"和"1142"被征用到link攻击里去了（这个只能通过反复多次听编号才能得知）。那么我需要用010Editor这个十六进制编辑器打开HIRC文件，打开后使用无符号整形(ui32)搜索并用“标签”记录整个Event组的wemID所有位置，也即数字序号为"408","786","975","1142"对应的wemID"308361897","571928584","696541074","798788669"在这个HIRC文件里的所有位置。我们会发现，这Event组4个成员的wemID在HIRC文件的3个地址块中紧凑出现，而唯独"786"和"1142"对应的wemID还另外在两个地方地址块紧凑出现。那我们只需要把单独出现的这两组"786"和"1142"对应的wemID替换成原本就只在link攻击组里别的成员的wemID，保存并得到新的HIRC文件即可。

Q10: 转化完新的bnk该怎么打包成mod并使用呢？
A10: 具体步骤如下：
1. 去这个网站<http://nenkai.github.io/relink-modding/getting_started/mod_manager/>安装“Reloaded-II Mod Manager”和“gbfrelink.utility.manager”并确保它们都是最新版本。
2. 运行Reloaded-II这个mod管理器，找到左栏上面的“3个齿轮图标”，然后在右边点击“新增”按钮，然后往里面根据实际需要填mod的信息，填写完后你会发现在mod管理器Release文件夹里的“Mods”中找到刚刚创建的mod文件夹，然后根据你bnk|pck文件相对于“data”解包文件的具体位置，创建一系列文件夹再把bnk|pck文件丢进去即可。比如我的新bnk文件就放在<.\Release\Mods\gbfr.voice.Siegfried_DARKv1.1.2\gbfr.voice.Siegfried_DARK\GBFR\data\sound\Japanese\>这个路径下。
3. 在Reloaded-II左栏找到上面的“添加应用程序”，选择游戏根目录下的exe文件，然后在右边就会出现mod，点击前面的方框加号就能使用了。
'''

commands = [Action2Motion_Process, Motion2Event_Process, Event2wemID_Process, MixUp_Process, AltWavName_Process, bnkRepack_Process]

class TextRedirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, string):
        self.widget.insert(tk.END, string)
        self.widget.see(tk.END)

    def flush(self):
        pass

def add_to_text_box(message):
    text_box.insert(tk.END, message + '\n')
    text_box.see(tk.END)

def execute_command(button, light, process_function):
    light.config(bg='yellow')
    button_text = button['text']
    add_to_text_box(f'=====================================================Command {button_text} is running.....')
    try:
        process_function()
        light.config(bg='green')
        add_to_text_box(f'=====================================================Command {button_text} completed successfully.')
    except Exception as e:
        light.config(bg='red')
        add_to_text_box(f'=====================================================An error occurred with command {button_text}: {e}')
    finally:
        button['state'] = 'normal'

def update_language(value):
    global text_box
    for widget in root.winfo_children():
        widget.destroy()

    if value == 'English':
        title_text = title_text_Eng
        button_texts = button_texts_Eng
        First_text = First_text_Eng
        indicates = indicates_Eng
        qnatext = qnatext_Eng
    elif value == '简体中文':
        title_text = title_text_Ch
        button_texts = button_texts_Ch
        First_text = First_text_Ch
        indicates = indicates_Ch
        qnatext = qnatext_Ch

    text_label = tk.Label(root, text=First_text, font=('Helvetica', 10, 'bold'), bg='pink', justify=tk.LEFT)
    text_label.pack(fill='x')

    for i in range(len(commands)):
        frame = tk.Frame(root)
        frame.pack(fill='x', padx=3, pady=5)

        button_text = button_texts[i]
        button = tk.Button(frame, text=f'Command {i + 1}', width=33, height=1, font=('Helvetica', 10, 'bold'),
                           anchor='w')
        button.pack(side='left')

        light = tk.Label(frame, width=1, height=1, bg='grey')
        light.pack(side='left')

        lines = indicates[i].count('\n') + 1
        indicate = tk.Text(frame, fg='#0000FF', height=lines, width=125, bg='#F0F0F0', bd=0, relief='flat')
        indicate.insert(tk.END, indicates[i])
        indicate.pack(side='left', fill='x', padx=3)
        indicate.config(state='disabled')

        def make_callback(button, light, command):
            def callback():
                button['state'] = 'disabled'
                threading.Thread(target=execute_command, args=(button, light, command)).start()

            return callback

        button.config(command=make_callback(button, light, commands[i]), text=str(button_text))

    # 在text_box上方添加一个新的Frame
    top_frame = tk.Frame(root)
    top_frame.pack(fill='x')
    top_frame.pack(fill='x', padx=5)

    indicate = tk.Text(top_frame, fg='#000000', height=1, font=('Helvetica', 10, 'bold'), bg='#F0F0F0', bd=0, relief='flat')
    indicate.insert(tk.END, "========== OUTPUT MESSAGE ==========")
    indicate.pack(side='left', fill='x')
    indicate.config(state='disabled')

    language = tk.StringVar()
    language.set('# Language 语言 #')
    # 创建下拉菜单
    language_menu = tk.OptionMenu(top_frame, language, 'English', '简体中文', command=update_language)
    language_menu.pack(side='right')

    def open_dialog():
        dialog = tk.Toplevel(root)
        dialog.title('# Q&A 提问与回答')
        text_area = scrolledtext.ScrolledText(dialog, wrap=tk.WORD)
        text_area.insert(tk.INSERT, qnatext)
        text_area.config(state=tk.DISABLED)  # 设置为只读
        text_area.pack(fill=tk.BOTH, expand=True)  # 让组件随窗口大小变化

    qa_button = tk.Button(top_frame, text='# Q&A #', command=open_dialog)
    qa_button.pack(side='right')

    # 创建文本框
    text_box = ScrolledText(root, height=10, bg='#E0E0E0', bd=0, relief='flat')
    text_box.pack(fill='both', expand=True)

    # 重定向标准输出和错误到文本框
    sys.stdout = TextRedirector(text_box)
    sys.stderr = TextRedirector(text_box)

    root.title(title_text)

root = tk.Tk()
# root.configure(bg='light blue')

# 创建文本框
text_box = ScrolledText(root, height=10, bg='#E0E0E0', bd=0, relief='flat')
text_box.pack(fill='both', expand=True)

update_language('English')

root.mainloop()
