Here is the translation of the text into English:

This tool is designed for creating voice mods for the game "Granblue Fantasy: Relink". It conveniently extracts the relationships between each character's "Action", "Motion", "EventName", and the corresponding "wemID" in the bnk files, and generates an Excel table. Additionally, after filling out the table, it can repackage the content into the corresponding bnk files. If you have any questions, you can contact me on BiliBili by searching for Dangoooooo, or find me through QQ number: 1041271418. You can also find the source code and the newly released tool package directly on GitHub at <http://github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>. I will check periodically.

Please read the Q&A section below carefully, as well as the comments on the GBFR#EasyPickUpVoiceInfo.exe tool.

>>>Q&A<<<

Q1: Why do we need to find the corresponding relationships?
A1: Because the character voices in this game can be reused, multiple moves you execute through commands may correspond to the same set of voices, which can cause some trouble when creating voice mods.

Q2: What are the specific corresponding relationships?
A2: 
1. "Action" refers to the moves (the moves we command the characters to perform) ------ corresponding to multiple "Motions" (combinations of actions) [recorded in the plxxxx.msg files]
2. Each "Motion" corresponds to multiple "EventNames" (which include sounds, mouth movements, facial expressions, etc.) [recorded in the bxm series files in the pl\plxxxx folder, we only care about the _se.bxm files, which record the EventNames related to sound]
3. Each wem file has a corresponding ID and its sequence number (NumSerial) in the bnk file, and the EventName related to sound is engraved inside the wem file (it can be viewed and extracted individually by opening it in hexadecimal mode, such as with 010Editor). Note that the sound EventName recorded in the bxm file is slightly different from the one engraved in the wem file. Specifically, the EventName inside the wem file will have the last two digits of the four-digit character code after "pl" changed to something like 01, 02, etc., according to game updates, and there will be suffixes like _1, _2, or _a, _b at the end, while the EventName recorded in the bxm file is more like a summary. In summary, the corresponding format is wemID---NumSerial---EventName(with Serial) one-to-one correspondence, and multiple EventNames(with Serial) correspond to one EventName(in .bxm)
4. Multiple wems are packaged into multiple bnk or pck files according to the content and updates of the game. [In the bnk rules of GBFR, the vo_plxxxx.bnk file is used as the master HIRC file and contains no voice content. All others that have voice content are suffixed after the HIRC file with something like _01_01, _m, _town, or a combination of these]

Q3: What exactly do I need to fill out after generating the table?
A3: We only care about the "VoiceInfo_MixUp" table, and what needs to be filled out is columns ABCD in the "YourRecord" sheet of the table, where the second row contains instructions and the third row shows a demonstration. Specifically, column B is for writing the sequence number corresponding to the original audio ("belonging bnk|pck package"_"numerical identifier"), column C is for writing the name of the new audio wav file you want to replace it with (do not add the file extension). Column A is for naming the action corresponding to this wem file, which you can either fill in yourself, base on the move names I extracted from the GBFR logs plugin, or ignore altogether. Column D is for notes, completely based on your own needs. The remaining columns are automatically filled with built-in formulas.

Q4: How do I find the sequence number corresponding to the original audio?
A4: The letter code for the belonging bnk|pck package can be found in the instructions in the second row of column B, and there are two ways to know the numerical identifier:
1. You can directly consult the "MixUp" sheet in the "VoiceInfo_MixUp" table, find out which wem audios correspond to known Action numbers or names, or if there is no corresponding Action and Motion information, you can also infer the function of that audio through the specific content of "EventName" under the "PL_vo_Event_WithSerial" column.
2. You can also generate AI voice on websites (there should be many, I recommend a free website in China <https://ttsmaker.cn/>), let it read the numbers (such as a245, b431, etc.), then use the slicing tool <https://github.com/flutydeer/audio-slice> to cut it open, and use this tool software to package it back into the game. This way, you can intuitively hear which wem file your move corresponds to. However, the downside is that often a single action will correspond to multiple wem files, and you may need to repeat the action multiple times to hear it in full. I also provide Chinese and English versions of the pre-sliced numbering zip file "Serial(unpack to wav_org folder)" in the release section at <http://github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>. It ranges from letters a to j, each counting from 0 to 2000 (I will add some variations to the English letters to avoid confusion, which should be sufficient).

Q5: Why not make a one-click unpacking and repacking for pck files?
A5: Because the pck voice files involve loops in the voices, and you need to set the loop points yourself, and most of the battle voices are in the bnk files (except for the master control HIRC vo_plxxxx.bnk), the main reason is that I haven't found a tool online that can batch replace wem files in pck files, I'll do it when I have a chance. However, when I generated the comparison table, I also considered the pck files, and you can unpack and repack pck files with the "RingingBloom" software <https://github.com/Silvris/RingingBloom>.

Q6: What do I need to prepare in advance?
A6: 
1. First, you need to use "GBFRDataTools" <https://github.com/Nenkai/GBFRDataTools> to unpack the "data.i" file in the game's root directory <.\Steam\steamapps\common\Granblue Fantasy Relink>. You will get a huge "data" folder in the same directory (it is recommended to copy and rename the previous "data" folder to "data_org" as a backup) containing all the resources of the entire game, which is about the same size as the game itself, so you should reserve at least 70GB of hard disk space.
2. You need to have "Wwise Launcher" installed on your computer (download and install wwise from <https://www.audiokinetic.com>), which is used to help you convert wav files to wem files.

Q7: What if I want the character's lips to move when speaking lines after the replacement?
A7: This issue essentially involves engraving the correct EventName into the new wem.
The specific operation is:
1. Open Wwise Launch and create a new project, then in Wwise Launcher, select "Share Sets" > "Conversion Settings" > "Default Conversion Settings" in the upper middle part of the left column and double-click to open, find "Insert filename marker" on the right side, and check it.
2. In Wwise Launcher, select "Audio" > "Interactive Music Hierarchy" in the upper middle part of the left column, right-click in the "Default work unit" below and choose "Import Audio Files..." to add the "wav_AltName" audio folder generated by this tool software (#5) in the Yours folder, and directly right-click here to select "Convert..." to convert, which can convert all the renamed wavs to wems with EventName. The converted wem files are located under the project you created at <This PC\Document\WwiseProjects\(your project name)\.cache\Windows\SFX\>.

Certainly, here is the translation for Q8 to Q10:

Q8: Can I modify the correspondence between Motion and EventName in the bxm file?
A8: Of course, you can. The operation steps are as follows:
1. By completing step (#2) of this tool software, you can get the "Motion_EventName_ComparisonTable" comparison table in the "Yours" folder, and also get the series of files that have been unpacked from bxm to xml at <.\Yours\plxxxx\_se_bxm_xml_Group\>.
2. You can easily open the xml file with Notepad and find that it consists of many lines of "Seq LayerFlag" entries. I've excerpted one of the entries, which looks like this: <Seq LayerFlag="4294967295" StartTime="0.016667" SeqFlag="0" EventName="PL1100_vo_ATK_default_s" PartsNo="0" Flag="0" OffsetX="0.000000" OffsetY="0.000000" OffsetZ="0.000000"/>.
3. We can "delete" the entire line of entries, but we need to pay attention to modifying the "SeqNum" at the top of the xml file to make it equal to the current total number of entries.
4. We can also "replace" the "EventName" parameter in the entry with another existing "EventName" from another bxm file. It is very important to note that if you use the "replace" operation, then the duration of the group of wem audios corresponding to the new "EventName" must be shorter than any of the old wem group! (Generally, most of our wem groups are replacing the same voice anyway.) Otherwise, when a certain action A triggers the voice, it will prevent the voice of the next derivative action B from playing! If you need a long voice and it's hard to cut, there are two solutions: (a) modify the "StartTime" parameter in the entry to reduce it to 0, which can extend the playback time to some extent; (b) modify the "PartsNo" parameter in the entry, if its value is different from another bxm file's "PartsNo" value after the derivative action (usually the values are "0" and "1"), which means they belong to different "channels", i.e., the long voice triggered by action A will play simultaneously with the voice of the next derivative action B! This can also solve the problem of long voices to some extent.
5. Finally, you can pack the modified xml file into a bxm file using <.\Yours\nier_cli_mgrr\nier_cli_mgrr.exe> or by directly dragging the xml file into the exe.

Q9: Is it possible that some actions use individual audios from other Event groups? If so, what should I do?
A9: This situation does indeed occur. It may appear in the audios corresponding to link attacks. When this happens, we need to modify the master control HIRC file, which is the vo_plxxxx.bnk file, as it contains the calling rules. How to modify it can be explained through an example: When I was making the Siegfried voice mod, I found that "786" and "1142" from the "PL1100_vo_ATK_default_l" Event group, which has sequence numbers "408", "786", "975", "1142", were used in link attacks (this can only be known by repeatedly listening to the numbers). So, I needed to use the hexadecimal editor 010Editor to open the HIRC file, search for all the positions of the entire Event group's wemIDs using unsigned integer (ui32) search and "label" them, i.e., the sequence numbers "408", "786", "975", "1142" corresponding to wemIDs "308361897", "571928584", "696541074", "798788669" in the HIRC file. We will find that the wemIDs of these four members of the Event group appear compactly in three address blocks in the HIRC file, while only the wemIDs corresponding to "786" and "1142" appear compactly in two other places. We just need to replace the wemIDs corresponding to "786" and "1142" that appear separately with the wemIDs of other members that are only in the link attack group, save and get the new HIRC file.

Q10: How do I package the new bnk into a mod and use it?
A10: 
1. Go to this website <http://nenkai.github.io/relink-modding/getting_started/mod_manager/> to install "Reloaded-II Mod Manager" and "gbfrelink.utility.manager" and make sure they are the latest versions.
2. Run the Reloaded-II mod manager, find the "three gears icon" on the left column, then click the "Add" button on the right side, and fill in the mod information according to actual needs. After filling it out, you will find the newly created mod folder in the "Mods" inside the Release folder of the mod manager, then create a series of folders according to the specific location of your bnk|pck file relative to the "data" unpacked file, and throw the bnk|pck file into it. For example, my new bnk file is placed at <.\Release\Mods\gbfr.voice.Siegfried_DARKv1.1.2\gbfr.voice.Siegfried_DARK\GBFR\data\sound\Japanese\>.
3. In the Reloaded-II left column, find the "Add Application" at the top, select the exe file in the game's root directory, then the mod will appear on the right side, click the square plus sign in front to use it.

------------------------------------------------------------------------------------

这个工具是给《Granblue Fantasy: Relink》这个游戏做语音mod用的，它能很方便的提取出游戏里每个角色的“Action”、“Motion”、“EventName”、以及对应bnk文件里“wemID”的对应关系，并生成Excel表。此外也能在填写完表格后，根据其内容重新打包成对应bnk文件。如果有任何问题，都可以在BiliBili网站搜Dangoooooo联系我，也能通过qq号：1041271418找到我，也能直接在github上找到源代码和新发布的工具包 <http://github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>，我会不定时查看。

请仔细阅读下面的Q&A版块，以及GBFR#EasyPickUpVoiceInfo.exe工具上的注释。

>>>Q&A<<<

Q1: 为什么要找出对应关系？
A1: 因为这个游戏的角色语音会存在复用的情况，你发出指令搓出的多个招有可能会对应同一组语音，会给我们做语音mod带来不少困扰。

Q2: 对应关系具体是怎么样的呢？
A2: 
1. Action就是指招数（我们发出指令控制人物发出的招数）------对应多个Motion（动作组合）[ 记录在plxxxx.msg文件里 ]
2. 每个Motion又对应多个EventName（其中有声音、嘴巴动作、脸部表情等等）[ 记录在pl\plxxxx文件夹里的bxm系列文件，我们只关心_se.bxm文件，它里面记录着与声音相关的EventName ]
3. 每个wem文件都有对应的ID，以及它在bnk文件里的序列号（NumSerial），且和声音相关的EventName会刻在wem里文件内部（要用16进制方式打开方可单独查看和提取，比如010Editor）。要注意的是记录在bxm文件里的声音EventName和刻在wem里的略微不一样。具体表现是wem内的EventName在pl后的4位角色编号后2位会根据游戏更新变成像01、02这样，同时在末尾会出现_1、_2或者_a、_b这样的后缀，而bxm内记载的EventName更像一个汇总信息。综上所述，对应形式是wemID---NumSerial---EventName(with Serial)一一对应，而多个EventName(with Serial)会对应一个EventName(in .bxm)
4. 多个wem会根据游戏的内容和更新封装成多个bnk或者pck。[ GBFR的bnk制定规则里vo_plxxxx.bnk文件是用作总控的HIRC文件，没有语音内容。除此以外的都是具有语音内容的，都是在HIRC文件后面加类似于_01_01或者_m或者_town的后缀，亦或是它们的混合]

Q3: 生成表后我具体要填写什么呢？
A3: 我们只关心"VoiceInfo_MixUp"表，需要填写的是表里的"YourRecord"sheet内的ABCD列，其中第2行有说明，第3行有演示。具体是B列填写你想原音频所对应序号（"所属bnk|pck包"_"数字编号"），C列填写你想替换成的新音频的wav文件名（不要多填入文件扩展名）。而A列是你给这个wem名字对应的动作起的名字，你既可以删掉自己填，也能根据我从显血插件GBFR logs里提取出的招数名，还能直接忽视不填。D列是完全根据自己需求填写的备注。剩余列则内置公式自动填充。

Q4: 那我该怎么找到原音频所对应的序号呢？
A4: 所属bnk|pck包可在B列第2行的说明里找到字母代号，而后面的数字编号有两种办法得知：
1. 可以直接查阅"VoiceInfo_MixUp"表里的“MixUp”sheet内，通过已知Action编号或者名字来查找它们对应哪些wem音频，或者如果没有对应的Action和Motion等信息，也能通过“PL_vo_Event_WithSerial”列下的“EventName”的具体内容来推测该音频的作用。
2. 你也可以在网站（应该挺多的，我这边推荐一个中国内免费的网站<https://ttsmaker.cn/>）上生AI语音，让它读编号（比如读a245，b431等），再利用切片工具<http://github.com/flutydeer/audio-slice>切开，并利用这个工具软件封包回游戏内。你就能很直观的听到你的招数会对应哪个wem文件了。不过缺点是很多时候某个动作会对应多个wem文件，也许你需要重复多次这个动作才能听完整。我这边也在<http://github.com/Dangoooooooooo/GBFR_EasyPickUpVoiceInfo>的release里提供了中文版和英文版已经切片好的编号压缩包“Serial(unpack to wav_org folder)”，从a到j，每个从0数到2000（由于英文字母可能会听混，我会加一些变化上去，应该是够用了）。

Q5: 为什么不做pck的一键解包和封包？
A5: 因为pck语音文件内涉及到语音的循环什么的，要自己设置循环点，且绝大部分的战斗语音都在bnk文件里（除了vo_plxxxx.bnk这个总控HIRC的），最主要的是我没找到网上有能批量替换pck内wem文件的工具，有机会再弄了。不过我生成对照表的时候也考虑到了pck文件，你可以通过"RingingBloom" <http://github.com/Silvris/RingingBloom>这个软件去解包和封包pck文件。

Q6: 有什么需要事先准备的？
A6: 
1. 首先你需要用"GBFRDataTools" <http://github.com/Nenkai/GBFRDataTools> 解包在游戏根目录下<.\Steam\steamapps\common\Granblue Fantasy Relink>的 "data.i" 文件。你会在同目录下得到一个超大的"data"文件夹（推荐把之前的“data”文件夹复制改名备份成"data_org"）里面包含了整个游戏所有的资源，它的大小和游戏本体大小差不多，至少要预留70个G的硬盘空间。
2. 你的电脑里需要装有“Wwise Launcher”(从<http://www.audiokinetic.com>下载后安装wwise)，它是用来帮你把wav转化为wem文件。

Q7: 我希望替换完后人物的嘴唇要在念台词时张嘴怎么办？
A7: 这个问题本质上就是将正确的EventName刻在新的wem里。具体操作是：
1. 打开Wwise Launch新建一个项目，接着在Wwise Launcher里左栏靠中上部依次选择“Share Sets”>“Conversion Settings”>“Default Conversion Settings”并双击打开，将右边找到“Insert filename marker”，并勾选上即可。
2. 在Wwise Launcher里左栏靠中上部依次选择“Audio”>“Interactive Music Hierarchy”，在下面的“Default work unit”里右键选择“Import Audio Files...”添加本工具软件(#5)在Yours文件夹生成的“wav_AltName”音频文件夹，并直接在这里单击右键选择“Convert...”进行转化，即可把刚刚改好名的所有wav转为带有EventName的wem。转化完成的wem文件在<此电脑\Document\WwiseProjects\（你的项目名）\.cache\Windows\SFX\>你新建的项目下。

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
A10: 
1. 去这个网站<http://nenkai.github.io/relink-modding/getting_started/mod_manager/>安装“Reloaded-II Mod Manager”和“gbfrelink.utility.manager”并确保它们都是最新版本。
2. 运行Reloaded-II这个mod管理器，找到左栏上面的“3个齿轮图标”，然后在右边点击“新增”按钮，然后往里面根据实际需要填mod的信息，填写完后你会发现在mod管理器Release文件夹里的“Mods”中找到刚刚创建的mod文件夹，然后根据你bnk|pck文件相对于“data”解包文件的具体位置，创建一系列文件夹再把bnk|pck文件丢进去即可。比如我的新bnk文件就放在<.\Release\Mods\gbfr.voice.Siegfried_DARKv1.1.2\gbfr.voice.Siegfried_DARK\GBFR\data\sound\Japanese\>这个路径下。
3. 在Reloaded-II左栏找到上面的“添加应用程序”，选择游戏根目录下的exe文件，然后在右边就会出现mod，点击前面的方框加号就能使用了。
