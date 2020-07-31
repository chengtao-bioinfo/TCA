#!/usr/bin/perl -w
# 2020/7/15 v20200715 版本发布
# 2020/7/20 v20200720 版本发布
#           v20200720 版本更新时间：2020/7/20
#           v20200720 版本更新内容：详见

use strict;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtUnicode;
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Chart::Line;
use Encode;
use Win32;
use Win32::GUI();
use File::Path qw(remove_tree);

my ($mday,$mon,$year) = (localtime)[3..5];
$mday = sprintf("%d", $mday);
$mon  = sprintf("%d", $mon + 1);
$year = $year + 1900;

# my $dateXXX = sprintf ("%4d%02d%02d", $year,$mon,$mday);
# my $TrialLim = sprintf ("%d%d%d%d%d%d%d%d", ord('C')-65,ord('A')-65,ord('B')-65,ord('I')-65,ord('A')-65,ord('E')-65,ord('D')-65,ord('A')-65);
my $version = "STR报告自动生成软件 v0.1 (开发版)";

my $pwd = `cd`;
chomp $pwd; print "L28:" . $pwd . "\n";

my %sampleDate;  # 记录 "样本编号 (.PrevSamples.txt) / 实验编码 (供患信息)" 对应的 "采样日期"  【.PrevSamples.txt】
my %receiveDate;  # 记录 "样本编号 (.PrevSamples.txt) / 实验编码 (供患信息)" 对应的 "收样日期"  【.PrevSamples.txt】
my %sampleType;  # 记录 "样本编号 (.PrevSamples.txt) / 实验编码 (供患信息)" 对应的"样本类型" 存入 %sampleType  (********* 样本编号 对应 供患信息 中的 实验编码，前后最好统一起来 ***********)  【.PrevSamples.txt】
my %Chimerism;  # hash表，用于存储每个 患者编码 对应的每次检测的 嵌合率 结果  【.PrevSamples.txt】
my %SampleID;  # hash表，用于存储每个 患者编码 对应的每次检测的 实验编码 / 患者编码: HUN001胡琳  (构建每个患者的唯一编码, 医院编号+姓名)  【.PrevSamples.txt】
my %HasChimerism;  # 实验编码 与 患者编码的对应关系，如: STR1708646-T <=> HUN001胡琳  【.PrevSamples.txt】
my %ReportDate;  # hash表，用于存储每个 患者编码 对应的 每次检测 的 报告日期  【.PrevSamples.txt】
my @exp_list;  # 将 实验编码 存入数组 @exp_list  【供患信息】
my $i = 0;
my @data_in;  # 保存输入 供患信息.txt 中每个实验的完整信息，共21列信息  【供患信息】
my %exp_id;  # exp_id 哈希，存放每个实验编码的原始顺序  【供患信息】
my @TCA_id;  # 将 报告单编号 存入数组 @TCA_id  【供患信息】
my %exp_num;  # exp_num 一维哈希，保存每个 报告单编号 中包含几个 实验编码   【供患信息】
my %exp_seq;  # 存储 报告单编号 对应的3/2/1份 实验编码 的先后顺序 (按 "术前患者", "术前供者"， "术后患者"，"术后供者")   【供患信息】
my %exp_error;  # 记录报告单编号 是否存在 实验错误, 0:没问题; 1:有问题   【供患信息】
my %together;  # %together 二维哈希，第一维是哈希，键是 报告单编号，第二维是列表，保存每个 报告单编号 对应的 实验编码   【供患信息】
my %identity;  # 存储 报告单编号 与 患者编码的对应关系，如: TCA1710480 <=> HUN001胡琳   【供患信息】
my %history;  # $history{$str[8]} = $#{$Chimerism{$identity{$str[8]}}}+1;  # 获取该 患者 在该医院总的检测次数   【供患信息】
my $test = 0;
my $error;
my $InputLoaded = 0;  # 标记 供患信息 是否读入： 1 已读入 / 0 未读入
my $SummaryLoaded = 0;  # 标记 已有数据 是否读入： 1 已读入 / 0 未读入
my $ExpLoaded = 0;  # 标记 下机数据 是否读入： 1 已读入 / 0 未读入

# 1省份   2医院编码   3医院全称   4医院别名1  5医院别名2  ...
# 示例如下：
#  安徽	AH002	安徽医科大学第一附属医院	安医附院	安医附一   ########
my %region;         # 医院全称 及 对应省份,例如： 安徽医科大学第一附属医院 <=> 安徽   【.HospitalTrans.txt】
my %ID;             # 医院全称 及 对应的编号,例如：安徽医科大学第一附属医院 <=> AH002   【.HospitalTrans.txt】
my %alias;          # 别名 及其对应的 医院全称,例如：安医附院 <=> 安徽医科大学第一附属医院   【.HospitalTrans.txt】
                    #                             安医附一 <=> 安徽医科大学第一附属医院   【.HospitalTrans.txt】

my $DOS = Win32::GUI::GetPerlWindow();
Win32::GUI::Hide($DOS);  # 隐藏后台运行的  console window  ** 测试时考虑不隐藏，一边检查 log **

# if ($dateXXX > $TrialLim){
        # $error = '测试版本，试用期至'.$TrialLim.'
# 请联系ct@gzjrkbio.com';
        # Win32::MsgBox ($error, 0, "已过期");
        # exit(0);
# }

# $error = '此版本可能存在错误，仅供测试使用
# 任何问题和建议请联系ct@gzjrkbio.com
# 是否继续？';

# my $goon = Win32::MsgBox ($error, 4, "声明");

# exit(0) if $goon == 7;

my %allele;  # 将 已有数据 (型别汇总) 及 下机数据 中 每个 实验编码 的每个marker 对应的型别信息，存入 %allele  【已有数据(型别汇总)】
my %PrevAllele;  # 保存 已有数据 (型别汇总) 每个 实验编码 的每个marker 对应的 型别信息  【已有数据(型别汇总)】
my $curr_index;

my %ThisAllele;  # 存放 下机数据 中该 实验编码 该marker 对应的型别信息  【下机数据】
my %area;  # 存放 下机数据 中该 实验编码 该marker 对应的 area信息
my %trans; # 用来保存 下机数据 中该 实验编码 缩写 到 全称 的转换

# 定义STR检测的位点编号
# my @markers = ('D8S1179','D21S11','D7S820','CSF1PO','D3S1358','D5S818','D13S317','D16S539','D2S1338','D19S433','VWA','D12S391','D18S51','Amel','D6S1043','FGA');  # 16 markers from DSBK
# 暂时使用 16 marker 进行测试
my @markers_jrk = ('D3S1358','VWA','D16S539','CSF1PO','TPOX','Yindel','Amel','D8S1179','D21S11','D18S51','Penta E','D2S441','D19S433','TH01','FGA','D22S1045','D5S818','D13S317','D7S820','D6S1043','D10S1248','D1S1656','D12S391','D2S1338','Penta D');
my %markerExist;  # 使用hash表 存放 marker名， 以方便方便检查 marker是否存在
foreach (@markers_jrk){
        $markerExist{$_} = 'yes',
}

# 定义供患信息表格的表头
my @headers = ('收样日期', '生产时间', '移植日期', '样品类型', '样品性质', '分选类别', '采样日期', '实验编码', '报告单编号', '姓名', '供患关系', '性别', '年龄', '诊断', '亲缘关系', '关联样本编号', '医院编码', '送检医院', '送检医生', '住院号', '床号');

my $InputSample_str = "未选择";  # 获取选中文件的文件名 (见L568, $tmpfilename)  【List1_DblClick 函数: "供患信息" 列表框中文件双击时调用的处理函数】
my $PrevExp_str = "未选择";  # 获取选中文件的文件名字  【List2_DblClick 函数: "已有数据" 部分的列表框中双击时的处理函数】
my $Output_Dir = "未选择";  # 获取生成报告的文件夹  【RUN_Click 函数: "生成报告" 按钮 点击 时的处理函数】
my $Output_rpt_str = "未选择";  # 生成报告文件的名称

my @InputFound;  # 表头合格的文件，文件名存入数组 @InputFound  【供患信息】
my @PrevFound;  # 表头合格的文件，文件名存入数组 @PrevFound  【已有数据(型别汇总)】
my @InputList;  # 使用 Shorten 函数，将文件名截断后，存入数组 @InputList  【供患信息】
my @PrevList;  # 使用 Shorten 函数，将文件名截断后，存入数组 @PrevList  【已有数据(型别汇总)】
my @ThisFound;  # 将选中的 下机数据.txt 文件名 存入 @ThisFound  【下机数据】
my @ThisList;  # 调用 Shorten函数截短输入文件名，同时 存入 @ThisList  【下机数据】

################# 配置文件读取部分 ###############################
# 定义配置文件 TCAconfig.ini 中关键字名称
my @ConfigList = ("InputLoc", "SummaryLoc", "ThisLoc", "OutputLoc");
my %ConfigHash;

# 设置 "InputLoc", "SummaryLoc", "ThisLoc", "OutputLoc" 默认值为 $pwd, 见Line27-28
foreach (@ConfigList){ my $tmp = $_ ; print "L113:" . $tmp . "\n";
                $ConfigHash{$tmp} = $pwd; print "L114:" . $ConfigHash{$tmp} ."\n";
}

&ReadConfig;  # 读取 TCAconfig.ini 文件
################# 配置文件读取完成 ###############################

################# UI部分设置开始 ###############################
################## 设置软件的主界面 ############################
my $main = Win32::GUI::Window->new(
        -name => 'Main',
        -text => $version,
        -width => 570,
        -height => 500,
        -pos => [200, 200],
        -sizabke => 0,
        -resizable => 0,
);

my $font = Win32::GUI::Font->new(
        -size => 24,
        -color => 0x0000FF,
);

my $sb = $main->AddStatusBar();  # 添加状态条?

################## 定义软件界面上的 供患信息 部分 ##########################
my $text11 = $main->AddLabel(  # 定义 "供患信息" 名称
        -text => '供患信息',
        -pos => [10, 10],
        -font => $font,
);

my $text12 = $main->AddLabel(  # 定义 "供患信息" 下方的 "已找到"
        -text => '已找到',
        -pos => [10, 45],
);

my $Input1 = $main-> AddListbox(  # 定义 "供患信息" 部分的列表框
        -name => "List1",
        -left => 10,
        -top => 60,
        -width => 250,
        -height => 80,
        -hscroll => 1,
        -vscroll => 1,
);

my $open1 = $main->AddButton(  # 定义 "供患信息" 部分列表框下方的 "其他位置" 按钮
        -name => "Open1",
        -text => "其他位置...",
        -pos  => [ 10, 130 ],
);

my $text14 = $main->AddLabel(  # 定义 "供患信息" 部分的 "双击读取" 提示信息
        -pos => [140, 135],
        -text => '↑双击读取',
);

my $text13 = $main->AddLabel(  # 定义 "供患信息" 部分的 "尚未读取" 提示信息
        -pos => [10, 155],
        -width => 250,
        -text => '尚未读取',
);

################## 定义软件界面上的 已有数据 部分 ##########################
my $text21 = $main->AddLabel(  # 定义 "已有数据" 名称
        -text => '已有数据',
        -pos => [300, 10],
        -font => $font,
);

my $text22 = $main->AddLabel(  # 定义 "已有数据" 名称下方的 "已找到"提示
        -text => '已找到',
        -pos => [300, 45],
);

my $Input2 = $main-> AddListbox(  # 定义 "已有数据" 部分的列表框
        -name => "List2",
        -left => 300,
        -top => 60,
        -width => 250,
        -height => 80,
        -hscroll => 1,
        -vscroll => 1,
);

my $open2 = $main->AddButton(  # 定义 "已有数据" 部分列表框下方的 "其他位置" 按钮
        -name => "Open2",
        -text => "其他位置...",
        -pos  => [ 300, 130 ],
);

my $text24 = $main->AddLabel(  # 定义 "已有数据" 部分列表框下方的 "双击读取" 提示信息
        -pos => [430, 135],
        -text => '↑双击读取',
);

my $text23 = $main->AddLabel(  # 定义 "已有数据" 部分列表框下方的 "尚未读取" 提示信息
        -pos => [300, 155],
        -width => 250,
        -text => '尚未读取',
);

################## 定义软件界面上的 打印已有分型 部分 ##########################
my $display2 = $main->AddButton(  # 定义 "打印已有分型" 按钮
        -name => "DISPLAY2",
        -text => "打印已有分型",
        -pos  => [ 10, 180 ],
        -size => [ 545 , 30],
        -disabled => 1,  # 设置默认为不可点击，仅在供患信息 及 已有数据读取完成后，变为可点击状态
);

################## 定义软件界面上的 添加下机数据 部分 ##########################
my $open3 = $main->AddButton(  # 定义 "添加下机数据" 按钮
        -name => "Open3",
        -text => "添加下机数据",
        -pos  => [ 10, 220 ],
        -size => [ 100, 30 ],
        -disabled => 1,  # 设置默认为不可点击，仅在打印已有分型完成后，变为可点击
);

my $del3 = $main->AddButton(  # 定义 "移除" 按钮
        -name => "Del3",
        -text => "移除",
        -pos  => [ 80, 250 ],
        -size => [ 30, 20 ],
        -disabled => 1,  # 设置默认为不可点击，仅在 "添加下机数据"? 完成后，变为可点击
);

my $Read3 = $main->AddButton(  # 定义 "读取" 按钮
        -name => "Read3",
        -text => "读取",
        -pos  => [ 500, 221 ],
        -size => [ 50, 50 ],
        -disabled => 1,  # 设置默认为不可点击，仅在 "添加下机数据"? 完成后，变为可点击
);

my $Input3 = $main-> AddListbox(  # 定义 "添加下机数据" 部分的列表框
        -name => "List3",
        -left => 120,
        -top => 220,
        -width => 370,
        -height => 60,
        -hscroll => 1,
        -vscroll => 1,
        -multisel => 2,
        -disabled => 1,
);

my $text3 = $main->AddLabel(  # 定义 "添加下机数据" 列表框下方"尚未读取"提示信息
        -text => "尚未读取",
        -pos => [120, 280],
);

################## 定义软件界面上的 添加下机数据 下方的分隔线 部分 ##########################
my $sep = $main-> AddLabel(  # 在"添加下机数据"部分下方添加分隔线
        -text =>"==============================================================================================================================",
        -pos => [0,300],
);

################## 定义软件界面上的 生成报告 部分 ##########################
my $run4 = $main->AddButton(  # 定义 "生成报告" 按钮
        -name => "RUN",
        -text => "生成报告",
        -font => $font,
        -pos  => [ 20, 320 ],
        -size => [160,60],
        -disabled => 1,  # 设置默认为不可点击，仅在 "添加下机数据" 完成后，变为可点击
);

my $RptBox = $main-> AddTextfield(  # 定义 "生成报告" 部分的文本框
        -name => "RptBox",
        -pos => [200, 310],
        -size => [350, 135],
        -readonly => 1,  # 设置为只读
        -multiline => 1,
        -vscroll => 1,
        -autovscroll => 1,
        -autohscroll => 0,
);

my $RUNwindow = new Win32::GUI::Window (  # 定义点击"生成报告"按钮后出现的新的提示界面
        -name  => "RUNWindow",
        -title => "正在生成文件，请稍候...",
        -pos   => [ 300, 300 ],
        -size  => [ 300, 100 ],
        -parent => $main,
        -sizabke => 0,
        -resizable => 0,
);

my $quit = $main->AddButton(  # 定义 "退出" 按钮
        -name => "QUIT",
        -text => "退出",
        -pos  => [ 20, 420 ],
        -size => [ 60,20],
);

my $copybutton = $main -> AddButton(  # 定义 "复制" 按钮
        -name => "COPY",
        -text => "复制",
        -pos  => [120, 420],
        -size => [ 60,20],
);

my $direct = 1;  ## 意义未名
$main->Show();  # 显示软件界面
################# UI部分设置完成 ###############################

################# 读取医院信息 .HospitalTrans.txt ##############
$sb -> Text('正在读取医院信息');
unless (open IN,".HospitalTrans.txt"){
        $error = "未找到医院信息
请检查或重新解压安装包！";
        my $s = Win32::MsgBox $error,1, "错误！";
        exit(0);
}
<IN>;
# .HospitalTrans.txt 的格式如下 ###################################
# 1省份   2医院编码   3医院全称   4医院别名1  5医院别名2  ...
# 示例如下：
#  安徽	AH002	安徽医科大学第一附属医院	安医附院	安医附一   ########
# 特别说明： 1) 不带表头; 2)字段间以TAB分隔 ##########################
while(<IN>){
        my @str = split;
        $region{$str[2]} = $str[0];  # 将医院全称及对应省份，保存到 %region
        $ID{$str[2]} = $str[1];  # 将医院全称及对应的编号，保存到 %ID

        while ($#str > 2){  # $#str:读取 @str最后一个元素的索引。等价于 @str >3。如果医院存在别名信息，则读取并保存到 %alias
          #### 遍历读取所有医院的别名，并将别名对应的医院全称对应关系，保存到 %alias
                my $tmp = pop @str;  # 从最右边的别名开始读取
                $alias{$tmp} = $str[2];  # 将别名 及其对应的 医院全称，保存到 %alias
        }
}
close IN;
################# 读取医院信息完成 ################################

################# 读取已有样本信息 .PrevSamples.txt ###############
$sb -> Text('正在检查必需文件');
unless (-e ".PrevSamples.txt"){
        $error = "未找到已有样本信息
请检查或重新解压安装包！";
        my $s = Win32::MsgBox $error,1, "错误！";
        exit(0);
}
open IN,".PrevSamples.txt";
<IN>;

# .PrevSamples.txt 的格式如下 ##############################################################################################################################################################################################################
# 0区域 1快递单号 2 3 4采样日期 5收样日期 6移植日期 7样品数量 8样品类型 9样品性质 10样本编号 11报告单编号 12姓名 13供患关系 14性别 15年龄 16诊断 17亲缘关系 18关联样本编号 19医院编码 20送检医院 21送检医生 22邮寄报告地址 23邮寄报告地址
# 示例如下：
#  区域	快递单号			采样日期	检测日期	移植日期	样品数量	样品类型	样品性质	样本编号	报告单编号	姓名	供患关系	性别	年龄	诊断	亲缘关系	关联样本编号	医院编码	送检医院	送检医生	邮寄报告地址	邮寄报告地址	月份	合并数据	样品类型	状态
# 湖南	不详				2017/7/1		[1 管  2 毫升 ]	骨髓-T细胞分选	术后	STR1708646-T	TCA1710480	胡琳	患者	女	17		胡琳	本人	HUN001	中南大学湘雅医院				7	HUN001胡琳术后	骨髓-T细胞分选	术后
# 湖南	不详				2017/7/1		[1 管  2 毫升 ]	骨髓-T细胞分选	术后	STR1708647-T	TCA1710481	周长新	患者	男	42		周长新	本人	HUN001	中南大学湘雅医院				7	HUN001周长新术后	骨髓-T细胞分选	术后
# 甘肃	8.24E+11			2017/6/29	2017/7/1		[2 管 2毫升]	全血	术后	STR1708648	TCA1710482	尹进龙	患者	男	13		尹进龙	本人	GS002	兰州军区总院	郭医生			7	GS002尹进龙术后	全血	术后
# 特别说明： 1) 带表头; 2) 快递单号 与 采样日期间空2列; 3)字段间以TAB分隔 ########################################################################################################################################################################
while(<IN>){
        chomp;
        my @str = split /\t/, $_;
        next unless $str[10];  # 跳过"样本编号"为空的行
        next unless $str[11];  # 跳过"报告单编号"为空的行
        next unless $str[12];  # 跳过"姓名"为空的行
        next unless $str[19];  # 跳过"医院编码"为空的行

        my $Smplid = $str[10];  # 读取"样本编号"
        $sampleDate{$Smplid} = $str[4]? $str[4]:'不详';  # 读取"采样日期"，若为空，则设置为"不详"，并将"样本编号"对应的"采样日期"存入 %sampleDate
        $sampleDate{$Smplid} = DateUnify($sampleDate{$Smplid});  # 使用DateUnify函数，调整日期格式
        $receiveDate{$Smplid} = $str[5]? $str[5]:'不详';  # 读取"收样日期"，若为空，则设置为"不详"，并将"样本编号"对应的"收样日期"存入 %receiveDate
        $receiveDate{$Smplid} = DateUnify($receiveDate{$Smplid});  # 使用DateUnify函数，调整日期格式
        $sampleType{$Smplid} = $str[8];  # 读取"样本类型"，并将"样本编号"对应的"样本类型"存入 %sampleType  （********* 样本编号 对应 供患信息 中的 实验编码，前后最好统一起来 ***********）
}

# print "STR1610506 Sample: ", $sampleDate{STR1610506},"\n";
# print "STR1610506 Recieve:", $receiveDate{STR1610506},"\n";

close IN;
################# 读取已有样本信息完成 ################################

################# 读取已有样本嵌合率结果 .PrevChimerism.txt ###############
unless (-e ".PrevChimerism.txt"){
        $error = "未找到已有嵌合率信息
请检查或重新解压安装包！";
        my $s = Win32::MsgBox $error,1, "错误！";
        exit(0);
}

open IN,".PrevChimerism.txt";
<IN>;

# .PrevChimerism.txt (术后结果汇总) 的格式如下 ##############################################################################################################################################################################################################
# 0报告编号        1患者姓名        2实验编码        3相关供者/报告        4嵌合率        5报告日期        6医院编号        7医院全称        8备注        9样本类型        10样品性质
# 示例如下：
# 报告编号	患者姓名	实验编码	相关供者/报告	嵌合率	报告日期	医院编号	医院全称	备注	样本类型	样品性质		检测次数
# TCA1710480	胡琳	STR1708646-T		99.73%		HUN001	中南大学湘雅医院		T细胞分选	术后	HUN001胡琳术后	14
# TCA1710500	林弦	STR1708661-T		71.04%		GD004	广州市第一人民医院		T细胞分选	术后	GD004林弦术后	9
# TCA1710501	林弦	STR1708661		93.56%		GD004	广州市第一人民医院		骨髓血	术后	GD004林弦术后	9
# TCA1710502	俞歆悦	STR1708571		99.71%		ZJ002	浙江省儿童医院		骨髓血	术后	ZJ002俞歆悦术后	10
# TCA1710503	蔡娇珍	STR1708677-B		81.02%		GD024	广东医科大学附属医院		B细胞分选	术后	GD024蔡娇珍术后	24
# TCA1710504	蔡娇珍	STR1708675		81.60%		GD024	广东医科大学附属医院		全血	术后	GD024蔡娇珍术后	24
# TCA1710505	蔡娇珍	STR1708676-T		86.03%		GD024	广东医科大学附属医院		T细胞分选	术后	GD024蔡娇珍术后	24
# TCA1710506	蔡娇珍	STR1708678-NK		77.20%		GD024	广东医科大学附属医院		NK细胞分选	术后	GD024蔡娇珍术后	24
# 特别说明： 1)带表头; 2)样本性质与检测次数间有一列无列名字; 3)字段间以TAB分隔 ########################################################################################################################################################################
while (<IN>){
        chomp;
        my @str = split /\t/, $_;
        next unless $str[4] =~ /\d+(\.\d+)?%/;  # 跳过嵌合率不为百分比格式的行（包括不存在或格式不对的行）
        next unless $str[2];  # 跳过 "样本编号" 不存在的行
        next unless $str[1];  # 跳过 "患者姓名" 不存在的行
        next if $str[7] =~ /N\/A/;  # 跳过 "医院全称" 为 N/A 的行
        next unless $str[7];  # 跳过 "医院全称" 不存在的行
        next if $str[7] eq '作废';  # 跳过 "医院全称" = '作废' 的行
        next unless $str[6];  # 跳过 "医院编号" 不存在的行

        # 判断 "医院全称" 对应的 "医院编号" 在 已有医院信息(.HospitalTrans.txt) 中是否存在
        if (exists $ID{$str[7]}){
                $str[6] = $ID{$str[7]};  # 存在, 根据已有医院信息，重新设置 "医院编号"
        }elsif(exists $alias{$str[7]}){  # 判断 "医院全称" （为别名）对应的 "医院全称" 在 已有医院信息(.HospitalTrans.txt) 中是否存在
                my $tmp = $alias{$str[7]};  # 存在，则获取 "医院别名" 对应的 "医院全称"
                $str[6] = $ID{$tmp};  # 根据已有医院信息，重新设置 "医院编号"
                $str[7] = $tmp;  # 根据已有医院信息，重新设置 "医院全称"
        }

        my $tmp = $str[6].$str[1];  # 如: HUN001胡琳  (构建每个患者的唯一编码, 医院编号+姓名)
        push @{$Chimerism{$tmp}}, $str[4];  # 数组里每一个元素都是 hash表，用于存储每个 患者编码 对应的每次检测的 嵌合率 结果
        push @{$SampleID{$tmp}}, $str[2];  # 数组里每一个元素都是 hash表，用于存储每个 患者编码 对应的每次检测的 实验编码
        $HasChimerism{$str[2]} = $tmp;  # 实验编码 与 患者编码的对应关系，如: STR1708646-T <=> HUN001胡琳

        # 判断 "报告日期" 是否存在
        if ($str[5]){  # 存在
                push @{$ReportDate{$tmp}}, DateUnify($str[5]);  # 数组里每一个元素都是 hash表，用于存储每个患者编码对应的每次检测的报告日期。报告日期 使用 DateUnify函数处理。
        }else{  # 不存在
                push @{$ReportDate{$tmp}}, "不详";  # 数组里每一个元素都是 hash表，用于存储每个患者编码对应的每次检测的报告日期。报告日期 不存在，则该次检测的报告日期存储为 "不详"
        }
}
close IN;
################# 读取已有样本嵌合率结果完成 ###############

#############################################################
################## 测试示例 ##########################
# my $teststr = "FJ001刘亦杰";
# $i = 0;
# foreach (@{$Chimerism{$teststr}}){
        # print "嵌合率：  ",$_,"\n";
        # print "样本编号：",${$SampleID{$teststr}}[$i],"\n";
        # print "报告时间：",${$ReportDate{$teststr}}[$i],"\n";
        # $i++;
# }

# foreach (keys %Chimerism){
        # print $_,"|", $#{$Chimerism{$_}}+1,"\n";
        # my $i = 0;
        # foreach my $tmp(@{$Chimerism{$_}}){
                # my $rptDate;
                # if ($ReportDate{$_}[$i] eq "不详"){
                        # $rptDate = sprintf "%s%d%s", "术后", $i+1 , "次";
                # }else{
                        # $rptDate = $ReportDate{$_}[$i];
                # }
                # printf "%15s%15s%20s\n", $rptDate, $SampleID{$_}[$i], $tmp;
                # $i ++;
        # }
# }
################## 测试示例 ##########################
#############################################################

################# 读取 InputLoc 目录下的 供患信息 文件 （只读取保存文件名） ###############
#Looking for InputSample files
my $temp = $ConfigHash{InputLoc};  # 读取供患关系信息存放路径，来源为 TCAconfig.ini, InputLoc
my @filelist = `dir /b $temp\\*.txt`; print "L476:" .$temp . "\n"; print "L476:" .@filelist. "\n";  # 读取 InputLoc下所有 txt文件，存储在 @filelist
my $localnumber = 1;

# 供患关系文件 的格式如下 ##############################################################################################################################################################################################################
# 0收样日期	1生产时间	2移植日期	3样品类型	4样品性质	5分选类别	6采样日期	7实验编码	8报告单编号	9姓名	10供患关系	11性别	12年龄	13诊断	14亲缘关系	15关联样本编号	16医院编码	17送检医院	18送检医生  19住院号   20床号
# 示例如下：
# 收样日期	生产时间	移植日期	样品类型	样品性质	分选类别	采样日期	实验编码	报告单编号	姓名	供患关系	性别	年龄	诊断	亲缘关系	关联样本编号	医院编码	送检医院	送检医生  住院号   床号
# 				术前			D19STR00039	QC-Q019	Q17	患者
# 				术前			10751	QC-Q019		供者
# 				术后			Q19	QC-Q019	Q19	患者							南京市儿童医院
# 	2020/3/5		[其他]	术前		2020/3/3	D20STR01231	TCA2007498	吴久芳	患者	男	47	-	吴久芳	本人
# 	2020/3/5		[其他]	术前		2020/3/3	D20STR01232	TCA2007498	吴文方	供者	男	44	-	弟弟
# 2020/3/4	2020/3/5		骨髓血	术后		2020/3/3	D20STR01230	TCA2007498	吴久芳	患者	男	47	-	吴久芳	本人		广东省人民医院	黄励思
# 	2018/6/8		全血	术前			STR1808282	TCA2007647	杨梅月	患者	女	不详		杨梅月	本人
# 	2018/6/8		全血	术前			STR1810793	TCA2007647	杨梅	供者	女	不详		姐姐
# 特别说明： 1)带表头; 2)字段间以TAB分隔; 3) 文件扩展名需要 ".txt" #################################################################################################################################################################################################
# 遍历每个txt文件
foreach (@filelist){
        chomp;
        my $tmpfilename = $temp."\\".$_; print "L481:" .$tmpfilename . "\n";  # 获得txt文件对应的完整路径名
        $sb -> Text("正在读取本地文件".substr('.....', 0,($localnumber++)%5+1));  # 输出状态信息
        next if /^\./;  # 跳过 InputLoc 下的 .
        (open IN ,$tmpfilename) || next;  # 打开文件，打不开则跳到下一个文件
        my $str = <IN>;  # 读取文件的表头行
        next unless $str;  # in case of empty file. 跳过空文件
        chomp $str;
        close IN;
        my $yes = 1;
        my @tmp = split /\t/, $str; print "L491:" . @tmp . "\n" ;
        # next if @tmp != 19;  # 若表头不是19列，则跳过该文件
        next if @tmp != 21;  # 若表头不是21列，则跳过该文件 （与List1_DblClick中格式保持一致）
        # 判断表头的前19个列名与 @headers 前19列是否完全一致，若存在不一致，则将 $yes = 0
        foreach my $i(0..20){ print "L493:".$tmp[$i]."\t".$headers[$i]."\n";  # 与List1_DblClick中格式保持一致，将 0..18 修改为 0..20
                $yes = 0 if $tmp[$i] ne $headers[$i];  # @headers = ('收样日期', '生产时间', '移植日期', '样品类型', '样品性质', '分选类别', '采样日期', '实验编码', '报告单编号', '姓名', '供患关系', '性别', '年龄', '诊断', '亲缘关系', '关联样本编号', '医院编码', '送检医院', '送检医生', '住院号', '床号');
        } print "L495:".$yes."\n";
        next if $yes != 1;  # 跳过表头列名 与 @headers 前19列不完全一致的文件
        push @InputFound, $tmpfilename;  # 表头合格的文件，文件名存入数组 @InputFound
        push @InputList, &Shorten($tmpfilename, 39);  # 使用 Shorten 函数，将文件名截断后，存入数组 @InputList
}
####
$Input1 -> Add(@InputList);  # 将@InputList中存储的截断文件名显示到 "供患信息" 部分的列表框中
################# 读取 InputLoc 目录下的 供患信息文件 完成 ###############

################# 读取 SummaryLoc 目录下的 已有数据 (已有型别汇总文件) （只读取保存文件名） ###############
# 已有数据文件（型别汇总） 的格式如下 ##############################################################################################################################################################################################################
# 0	1Marker1	2Marker2	3Marker3	4Marker4	5Marker5	6Marker6	7Marker7	8Marker8	9Marker9	10Marker10	11Marker11	12Marker12	13Marker13	14Marker14	15Marker15	16Marker16 ... 25Marker25
# 示例如下：
# 	D8S1179	D21S11	D7S820	CSF1PO	D3S1358	D5S818	D13S317	D16S539	D2S1338	D19S433	VWA	D12S391	D18S51	Amel	D6S1043	FGA
# DT2000011	13,16	29,32.2	8,11	11	16	10	11	12			14,19		13,14	X		21,23
# DT2000012	10,13	30	11	10,12	15,17	11	10,11	12			17,18		16	X,Y		22,25
# DT2000007	11,15	30,32.2	11,12	11,12	15,17	7,10	8	11,13			17,19		15,16	X,Y		21,24
# 特别说明： 1)带表头; 2)字段间以TAB分隔; 3) 文件扩展名需要 ".txt" #################################################################################################################################################################################################
#Looking for Previous Results files
$temp = $ConfigHash{SummaryLoc}; print "L505:" . $temp . "\n";  # 读取已有数据信息存放路径，来源为 TCAconfig.ini, SummaryLoc
@filelist = `dir /b $temp\\*.txt`;  # 读取 SummaryLoc 下所有 txt文件，存储在 @filelist
$localnumber = 1;
foreach (@filelist){
        chomp;
        my $tmpfilename = $temp."\\".$_; print $tmpfilename . "\n";  # 获得txt文件对应的完整路径名
        $sb -> Text("正在读取本地文件".substr('.....', 0,($localnumber++)%5+1));  # 输出状态信息
        next if /^\./;  # 跳过 SummaryLoc 下的 . (了解以下linux下的ll)
        (open IN ,$tmpfilename) or next;  # 打开文件，打不开则跳到下一个文件
        my $str = <IN>;  # 读取文件的表头行
        next unless $str;  # in case of empty file. 跳过空文件
        chomp $str;
        close IN;

        my $yes = 1;
        my @tmp = split /\t/, $str;
        next if @tmp != 26;  # 若表头超过17列，则跳过该文件  (需要修改为26，以支持JRK的25个marker)

        # 判断表头的列名 （marker名字）在 %markerExist 中是否存在，若有 %markerExist中没有的marker名，则将 $yes = 0
        # foreach my $i(1..16){  # 需要修改为 1..25，已支持JRK的25个marker；同时 %markerExist需要更新为25 marker
        foreach my $i(1..25){  # 需要修改为 1..25，已支持JRK的25个marker；同时 %markerExist需要更新为25 marker
                $yes = 0 unless exists $markerExist{$tmp[$i]};
        }
        next if $yes != 1;  # 表头中存在 %markerExist中没有的marker名，则跳过该文件
        push @PrevFound, $tmpfilename;  # 表头合格的文件，文件名存入数组 @PrevFound
        push @PrevList, &Shorten($tmpfilename, 39);  # 使用 Shorten 函数，将文件名截断后，存入数组 @PrevList
}
####
$Input2 -> Add(@PrevList);  # 将@InputList中存储的截断文件名显示到 "已有数据" 部分的列表框中
$sb -> Text('');  # 将状态信息清空
################# 读取 SummaryLoc 目录下的 已有型别汇总文件 完成 ###############

Win32::GUI::Dialog();  # 这句话是让窗口一直待机的一个循环，等待用户操作这个窗口
Win32::GUI::Show($DOS);  # my $DOS = Win32::GUI::GetPerlWindow();

exit(0);

#################################################################
# Main_Terminate 函数:主界面退出时, 调用 WriteConfig 函数           #
#################################################################
sub Main_Terminate {

        &WriteConfig;
        return -1;
}

#################################################################
# List1_DblClick 函数: "供患信息" 列表框中文件双击时调用的处理函数     #
#################################################################
sub List1_DblClick{
        if (@exp_list){

                $error = "已经成功读取数据，是否重新读取？";
                my $s = Win32::MsgBox $error,1, "注意！";
                # $Sure = 0;
                # $Msg1 -> DoModal();
                return 0 if $s != 1;
        }

        $InputLoaded = 0;
        $display2->Enable($InputLoaded*$SummaryLoaded);  # 在读取完 供患信息 文件名 及 已有数据 文件名后，将 "打印已有分型" 按钮，设置为"可点击"状态
        $open3->Enable($InputLoaded*$SummaryLoaded);  # 在读取完 供患信息 文件名 及 已有数据 文件名后，将 "添加下机数据" 按钮，设置为"可点击"状态

        $error = "尚未读取";
        $text13 -> Text($error);  # 将"供患信息" 部分的 "尚未读取" 提示信息设置为 "尚未读取"

        my $sel = $Input1->GetCurSel();  print "L667:" . $sel . "\n"; # 获取 "供患信息" 列表框中选中文件的下标
        $InputSample_str = $InputFound[$sel];  # 获取选中文件的文件名 (见L568, $tmpfilename)

        unless (open IN,$InputSample_str){  # 打开选中的文件
                $error = "文件打开失败！\n";
                Win32::MsgBox $error, 0, "错误！";
                return 0;
        }
        @exp_list = ();
        $i = 0;
        @data_in = ();  # 保存输入 供患信息.txt 中每个实验的完整信息，共21列信息
        %exp_id = ();
        @TCA_id = ();
        %exp_num = ();
        %exp_seq = ();
        %exp_error = ();
        %together = ();  # %together 二维哈希，第一维是哈希，键是 报告单编号，第二维是列表，保存每个 报告单编号 对应的 实验编码
        %history = ();
        $test = 0;

        my $tmp = <IN>;  # 读取表头行
        chomp $tmp;
        #if error content

        my $yes = 1;
        my @tmp = split /\t/, $tmp;
        $yes = 0 if @tmp != 21;  # 若表头行不是21列，则 $yes = 0
        foreach $i(0..20){
                $yes = 0 if $tmp[$i] ne $headers[$i]; # 判断表头列是否与 @headers 中一一对应
        }

        if ($yes != 1){  # 文件表头与 @headers 不是完全一致，报错
                $error = "这个文件貌似不对。表头应为：\n生产时间 移植日期 样品数量 样品类型 样品性质 分选类别 采样日期 实验编码\n报告单编号 姓名 供患关系 性别 年龄 诊断 亲缘关系 关联样本编号 医院编码\n送检医院 送检医生 住院号 床号\n";
                Win32::MsgBox $error, 0, "错误！";
                return;
        }

        # 遍历 供患信息.txt（前面已经读过表头行，这里从表头下一行开始读）
        #0生产时间   1移植日期   2样品数量   3样品类型   4样品性质   5分选类别   6采样日期   7实验编码   8报告单编号   9姓名   10供患关系   11性别   12年龄   13诊断   14亲缘关系   15关联样本编号   16医院编码   17送检医院   18送检医生  19住院号   20床号
        while (<IN>){
                chomp;
                my @str = split /\t/, $_;
                if ($str[7] eq ""){  # 实验编码 为空，报错
                        $error = "FATAL!! 实验编码为空！\n";
                        Win32::MsgBox $error, 0, "错误！";
                        exit(0);
                }
                $str[7] =~ s/\s+//g;  # 去除 实验编码 中的空字符
                if ($str[4] eq "" || $str[8] eq "" || $str[10] eq ""){  # 样品性质 / 报告单编号 / 供患关系 为空，报错，提示补充完整
                        $error =  "请补全样本信息（样品性质/报告单编号/供患关系）：".$str[7]."\n";
                        Win32::MsgBox $error, 0, "错误！";
                        exit(0);
                }
                $str[4] =~ s/\s+//g;  # 去除 样品性质 中的空字符
                $str[8] =~ s/\s+//g;  # 去除 报告单编号 中的空字符
                $str[10] =~ s/\s+//g;  # 去除 供患关系 中的空字符
                push @exp_list, $str[7];  # 将 实验编码 存入数组 @exp_list
                if (@TCA_id == 0 || $str[8] ne $TCA_id[-1]){  # 判断 @TCA_id 是否为空，或 报告单编号 不是 @TCA_id的最后一个
                        push @TCA_id, $str[8];  # 将 报告单编号 存入数组 @TCA_id
                }

                $exp_id{$str[7]} = $i; #exp_id 哈希，存放每个实验编码的原始顺序
                push @{$together{$str[8]}}, $str[7]; #together 二维哈希，第一维是哈希，键是 报告单编号，第二维是列表，保存每个 报告单编号 对应的 实验编码
                $exp_num{$str[8]} = @{$together{$str[8]}}; #exp_num 一维哈希，保存每个 报告单编号 中包含几个 实验编码

                ###检查新实验的数目####
                #0生产时间   1移植日期   2样品数量   3样品类型   4样品性质   5分选类别   6采样日期   7实验编码   8报告单编号   9姓名   10供患关系   11性别   12年龄   13诊断   14亲缘关系   15关联样本编号   16医院编码   17送检医院   18送检医生  19住院号   20床号
                if ($str[4] eq "术后" && $str[10] eq "患者"){

                        if (exists $HasChimerism{$str[7]}){  # 判断 %HashChimerism (.PrevChimerism.txt) 中是否已有该 实验编码 对应的结果记录. # 样本编号 与 患者编码的对应关系，如: STR1708646-T <=> HUN001胡琳
                                $error = "样本编号".$str[7]."(".$str[9].")已经存在嵌合率，本次将覆盖这条嵌合率信息";
                                ##注意是覆盖，所以后期要删掉这条信息
                                Win32::MsgBox $error, 0, "注意";
                                my $tmp = $HasChimerism{$str[7]};  print "L740:" . $tmp ."\n"; # 获取该 实验编码 对应的 患者编码
                                my $tmpNum;
                                foreach (0 .. $#{$Chimerism{$tmp}}){  # 遍历该 患者编码 对应的所有每次检测
                                        $tmpNum = $_ if ${$SampleID{$tmp}}[$_] eq $str[7];  # 获取当前次 实验编码 是 $SampleID{$tmp} 中的第几次实验
                                } print "L744:" . $tmpNum . "\n";
                                @{$Chimerism{$tmp}} = &DelItem(@{$Chimerism{$tmp}}, $tmpNum); # 删掉该 实验编码 在 %Chimerism 中的记录
                                @{$SampleID{$tmp}} = &DelItem(@{$SampleID{$tmp}}, $tmpNum);  # 删掉该 实验编码 在 %SampleID 中的记录
                                @{$ReportDate{$tmp}} = &DelItem(@{$ReportDate{$tmp}}, $tmpNum); # 删掉该 实验编码 在 %ReportDate 中的记录
                        }
                        $test ++;  # 记录实验次数
                        unless ($str[17]){  # 送检医院 为空
                                $history{$str[8]} = 0;  ##################### 待补充
                                $identity{$str[8]} = "NotFound";  ##################### 待补充
                        }else{
                                my $hospital = $str[17];  # 获取 送检医院
                                if (exists $ID{$hospital}){  # 假设 "送检医院" 为 "医院全称"，判断是否在 .HospitalTrans.txt 中存在
                                        $identity{$str[8]} = $ID{$hospital}.$str[9];  # 存储 报告单编号 <=> 患者编码(HUN001胡琳)
                                }elsif(exists $alias{$hospital}){  # 假设 "送检医院" 为 "医院简称"，判断是否在 .HospitalTrans.txt 中存在
                                        my $tmp = $alias{$hospital};  # 获取 "医院简称" 对应的 "医院全称"
                                        $str[17] = $tmp;  # 更新 "送检医院" 信息为 "医院全称"
                                        $identity{$str[8]} = $ID{$tmp}.$str[9];  # 存储 报告单编号 <=> 患者编码(HUN001胡琳)
                                }else{  # 送检医院 在 .HospitalTrans.txt 中不存在，报错，提示用户更新 .HospitalTrans.txt
                                        print $str[17],"木有找到\n";
                                        $error = $str[17]."木有找到请检查后添加到.HospitalTrans.txt";
                                        Win32::MsgBox $error, 0, "错误！";
                                        exit(0);
                                        $history{$str[8]} = 0;
                                        $identity{$str[8]} = "NotFound";
                                }
                        }

                        if ($identity{$str[8]} ne "NotFound"){  # 送检医院没问题(不为 空 && 在.HospitalTrans.txt中存在)
                                # 判断 患者编码 是否已经存在 嵌合率结果 (以前在该医院检测过)
                                if (exists $Chimerism{$identity{$str[8]}}){
                                        $history{$str[8]} = $#{$Chimerism{$identity{$str[8]}}}+1;  # 获取该 患者 在该医院总的检测次数
                                }else{
                                        $history{$str[8]} = 0;  # 该患者在该医院没有检测记录，第一次在这检测
                                }
                        }
                }else{  # 非 "术后" "患者" 样本，包括: "术前""患者" / "术前""供者"
                        # 判断 .PrevChimerism.txt 中是否已存在该 实验编码 的记录
                        if (exists $HasChimerism{$str[7]}){  # 若存在，则报错，显示提示信息
                                $error = "样本编号".$str[7]."(".$str[9].")显示存在嵌合率，但这并不是一个术后患者的样本！请检查后再次运行";
                                Win32::MsgBox $error, 0, "错误";
                                return -1;
                        }
                }
                ##############

                foreach my $tmp(0..20){
                        if ($str[$tmp]){
                                push @{$data_in[$i]}, $str[$tmp]; #data_in 二维列表，按InputSamples的顺序保存每一次实验信息
                        }else{  # 若该列信息为空，则设置为 "-"
                                push @{$data_in[$i]}, "-";
                        }
                }
                # print $str[7], "\t", $data_in[$i][14],"\n";
                $i ++;  # 更新实验编码序号
        }
        close IN;
        print "New Experiments: ",$test,"\n";
        $error = "已读取，新检测个数为：".$test;
        $text13 -> Text($error);  # 输出 供患信息.txt  读取结果

        # 遍历每个 报告单编号
        foreach my $TCAID(keys %together){  #together 二维哈希，第一维是哈希，键是 报告单编号，第二维是列表，保存每个 报告单编号 对应的 实验编码
                my @ddd=();  # 存储典型的初次检测3个样本的结果, $ddd[0] <=> 术前患者检测 实验编码; ddd[1] <=> 术前供者检测 实验编码; $ddd[2] <=> 术后患者检测 实验编码;
                my @temptogether = ();
                print "L808:",$TCAID,"|",$exp_num{$TCAID},"\n";  # 打印 报告单编号 对应的 实验编码 个数 (即，该报告对应的实验个数)
                $exp_error{$TCAID} = 0;  # 记录报告单编号 是否存在 实验错误, 0:没问题; 1:有问题
                if ($exp_num{$TCAID} == 3){  # 典型的初次检测，包括3个实验编号, 对应3份样本: "术前患者" + "术前供者" + "术后患者"
                        foreach my $exp_str_tmp(@{$together{$TCAID}}){
                                # print "$exp_str_tmp\n";
                                my $i = $exp_id{$exp_str_tmp};  # 获取该 实验编码 对应的是 供患信息.txt 中的第几个实验信息
                                if ($data_in[$i][4] eq "术前"){
                                        if ($data_in[$i][10] eq "患者"){
                                                $ddd[0] = $i;
                                        }else{  # "术前""供者"
                                                $ddd[1] = $i;
                                        }
                                }else{  # "术后"
                                        if ($data_in[$i][10] eq "患者"){
                                                $ddd[2] = $i;
                                        }else{
                                                $error = "报告编号：".$TCAID."不应包含术后供者样本，请检查！\n";
                                                Win32::MsgBox $error, 0, "错误！";
                                                $exp_error{$TCAID} = 1;  # 标记该 报告单编号对应的实验存在问题
                                        }
                                }
                        }
                        next if $exp_error{$TCAID} == 1;  # 跳过实验存在问题的报告单编号
                        $exp_seq{$TCAID} = join ",", @ddd;  # 存储该 报告单编号 对应的 3个实验编码, 以','分隔
                        # print $exp_seq{$TCAID},"\n";
                        # foreach my $i(0..2){
                                # print $ddd[$i],"|",$exp_list[$ddd[$i]],"|",$data_in[$ddd[$i]][4],"|",$data_in[$ddd[$i]][10],"\n";
                        # }
                }elsif ($exp_num{$TCAID} == 2){  # 报告对应2份 实验编码
                        my @total=();  # 存储两份实验编码的报告 顺序
                        foreach my $exp_str_tmp(@{$together{$TCAID}}){  # 遍历每个 报告单编号 对应的 实验编码
                                print "L840:$exp_str_tmp\n";
                                my $i = $exp_id{$exp_str_tmp};  # 获取该 实验编码 对应的是 供患信息.txt 中的第几个实验信息
                                my $sum = 0;
                                if ($data_in[$i][4] eq "术前"){
                                        $sum += 0;
                                }else{
                                        $sum += 2;
                                }
                                if ($data_in[$i][10] eq "患者"){
                                        $sum += 0;
                                }else{
                                        $sum += 1;
                                }
                                # "术前患者" $sum += 0
                                # "术前供者" $sum += 1
                                # "术后患者" $sum += 2
                                # "术后供者" $sum += 3
                                push @total, $sum; print "L857:$sum\n";
                                push @total, $i; print "L858:$i\n";
                        }
                        # 根据2个实验编码的sum值，确定二者的报告顺序？
                        if ($total[0]> $total[2]){  # 存储该 报告单编号 对应的2份实验编码 的先后顺序 （按 "术前患者", "术前供者"， "术后患者"，"术后供者"）
                                $exp_seq{$TCAID} = join ",", $total[3], $total[1];
                        }else{
                                $exp_seq{$TCAID} = join ",", $total[1], $total[3];
                        }
                        next if $exp_error{$TCAID} == 1;  # 没用
                        # print $exp_seq{$TCAID},"\n";
                        # my @tmpstr = split ",", $exp_seq{$TCAID};
                        # foreach my $i(@tmpstr){
                                # print $i,"|",$exp_list[$i],"|",$data_in[$i][4],"|",$data_in[$i][10],"\n";
                        # }
                }elsif ($exp_num{$TCAID} == 1){  # 该报告单编号 仅对应 1份实验编码，获取该实验编码在 供患信息.txt 中的行号，报错，输出提示信息
                        my $exp_str_tmp = ${$together{$TCAID}}[0];  # 获取该报告单编号 对应的 实验编码
                        print "L874:$exp_str_tmp\n";
                        my $i = $exp_id{$exp_str_tmp};  # 获取该 实验编码 对应的是 供患信息.txt 中的第几个实验信息
                        $error = "报告编号：".$TCAID."只包含一份实验样本，为".$data_in[$i][4].$data_in[$i][10]."样本。\n";
                        Win32::MsgBox $error, 0, "注意！";
                        $exp_seq{$TCAID} = $i;  # 后续关注此情况 %exp_seq 的使用情况
                        # print $i,"|",$exp_list[$i],"|",$data_in[$i][4],"|",$data_in[$i][10],"\n";
                }else{  # 报告单编号 对应的 实验编码 过多 （>3)，输出错误信息。但是并没有跳过该报告？
                        $error  = "报告编号：".$TCAID."的实验样本编号过多！请检查！\n";
                        Win32::MsgBox $error, 0, "注意！";
                        $exp_error{$TCAID} = 1;
                }
        }

        $InputLoaded = 1;
        $display2->Enable($InputLoaded*$SummaryLoaded);  # 将 "打印已有分型" 按钮,设置为可点击
        $open3->Enable($InputLoaded*$SummaryLoaded);  # 将 "添加下机数据" 按钮，设置为可点击状态
}

#################################################################
# List1_MouseMove 函数: "供患信息" 列表框中鼠标移上时的处理函数       #
#################################################################
sub List1_MouseMove{
        $sb -> Text('双击条目以读取');
}

#################################################################
# List1_MouseOut 函数: "供患信息" 列表框中鼠标移出时的处理函数        #
#################################################################
sub List1_MouseOut{
        $sb -> Text('');
}

###################################  2020.05.21.17.34 ###########################
################################################################################
# Open1_Click 函数: "供患信息" 部分列表框下方的 "其他位置" 按钮点击时的处理函数        #
################################################################################
sub Open1_Click{

        my @parms;
        push @parms,
          -filter =>
                [ 'TXT - Tab分隔文本', '*.txt'
                ],
          -directory => $ConfigHash{InputLoc},
          -title => '选择文件',
          -parent => $main,
          -owner => $main,;
        my @file = Win32::GUI::GetOpenFileName ( @parms );
        return 0 unless $file[0];
        # for (@file){
                # chomp;
                # print "$_\n";
        # }
        my $already = -1;
        foreach my $i(0 .. $#InputFound){
                my $tmp = $InputFound[$i];
                chomp $tmp;
                # print $tmp,"\n";
                # print $pwd.'\\'.$tmp,"\n";
                if ($tmp eq $file[0] or $file[0] eq $ConfigHash{InputLoc}.'\\'.$tmp){
                        $already = $i;
                        last;
                }
        }

        if ($already >= 0){
                $Input1 -> SetCurSel($already);
                return 0;
        }

        push @InputFound, $file[0];
        my $tmp = &Shorten($file[0], 37);
        $file[0] =~ /(^.+)\\[^\\]+$/;
        $ConfigHash{InputLoc} = $1;

        $Input1 -> InsertString($tmp);
        push @InputList, $tmp;

}

###################################################################################
# Open1_MouseMove 函数: "供患信息" 部分列表框下方的 "其他位置" 按钮鼠标移上时的处理函数    #
###################################################################################
sub Open1_MouseMove{
        $sb -> Text('从其他位置读取');
}

###################################################################################
# Open1_MouseOut 函数: "供患信息" 部分列表框下方的 "其他位置" 按钮鼠标移出时的处理函数     #
###################################################################################
sub Open1_MouseOut{
        $sb -> Text('');
}

###################################################################################
# List2_DblClick 函数: "已有数据" 部分的列表框中双击时的处理函数                        #
###################################################################################
################# 读取 SummaryLoc 目录下的 已有型别汇总文件 （只读取保存文件名） ###############
# 已有数据文件（型别汇总） 的格式如下 ##############################################################################################################################################################################################################
# 0	1Marker1	2Marker2	3Marker3	4Marker4	5Marker5	6Marker6	7Marker7	8Marker8	9Marker9	10Marker10	11Marker11	12Marker12	13Marker13	14Marker14	15Marker15	16Marker16 ... 25Marker25
# 示例如下：
# 	D8S1179	D21S11	D7S820	CSF1PO	D3S1358	D5S818	D13S317	D16S539	D2S1338	D19S433	VWA	D12S391	D18S51	Amel	D6S1043	FGA
# DT2000011	13,16	29,32.2	8,11	11	16	10	11	12			14,19		13,14	X		21,23
# DT2000012	10,13	30	11	10,12	15,17	11	10,11	12			17,18		16	X,Y		22,25
# DT2000007	11,15	30,32.2	11,12	11,12	15,17	7,10	8	11,13			17,19		15,16	X,Y		21,24
# 特别说明： 1)带表头; 2)字段间以TAB分隔; 3) 文件扩展名需要 ".txt" #################################################################################################################################################################################################
sub List2_DblClick{
        if (%PrevAllele){
                $error = "已经成功读取数据，是否重新读取？";
                my $s = Win32::MsgBox $error,1, "注意！";
                return 0 if $s != 1;
        }

        $SummaryLoaded = 0;
        $display2->Enable($InputLoaded*$SummaryLoaded);  # 后续测试关注此行的作用
        $open3->Enable($InputLoaded*$SummaryLoaded);  # 后续测试关注此行的作用

        $error = "尚未读取";
        $text23 -> Text($error);

        my $sel = $Input2->GetCurSel();  print "L984:" . $sel . "\n"; # 获取 "已有数据" 列表框中选中文件的下标
        $PrevExp_str = $PrevFound[$sel];  print "L985:" . $PrevExp_str . "\n" ;# 获取选中文件的文件名字

        %PrevAllele = ();

        unless ($InputLoaded){  # 需要先读取 "供患信息"
                $error = "请先读取供患信息！\n";
                Win32::MsgBox $error, 0, "错误！";
                return 0;
        }

        unless (open IN,$PrevExp_str){  # 打开 "已有数据" 双击选中的文件
                $error = "文件打开失败！\n";
                Win32::MsgBox $error, 0, "错误！";
                return 0;
        }
        my $tmp = <IN>;  # 读取 "已有数据" 选中文件的表头
        chomp $tmp;
        my $yes = 1;  # 用于记录选中文件的表头格式是否正确:1) 26列; 2) 25列marker名在 %markerExist中均存在
        my @tmp = split /\t/, $tmp;
        $yes = 0 if @tmp != 26;  # 判断 表头是否未 26列
        print "L1014:YES:",$yes,"\n";
        foreach $i(1..25){  # 判断表头中后25列对应的25个marker名 在 %markerExist 中是否存在，若有不存在的，则将 $yes = 0
                $yes = 0 unless exists $markerExist{$tmp[$i]};
                # print $tmp[$i],"|",$yes,"\n";
        }

        if ($yes != 1){  # 表头不对，报错，给出提示信息
                $error = "这个文件貌似不对\n";
                Win32::MsgBox $error, 0, "错误！";
                return;
        }
        print "L1025:" . $PrevAllele{STR154465}{D7S820},"\n";  ######### bug? %PrevAllele在前面只有定义，没有赋值
################################################  2020.05.22.08.10 ################################################
# 已有数据文件（型别汇总） 的格式如下 ##############################################################################################################################################################################################################
# 0	1Marker1	2Marker2	3Marker3	4Marker4	5Marker5	6Marker6	7Marker7	8Marker8	9Marker9	10Marker10	11Marker11	12Marker12	13Marker13	14Marker14	15Marker15	16Marker16 ... 25Marker25
# 示例如下：
# 	D8S1179	D21S11	D7S820	CSF1PO	D3S1358	D5S818	D13S317	D16S539	D2S1338	D19S433	VWA	D12S391	D18S51	Amel	D6S1043	FGA
# DT2000011	13,16	29,32.2	8,11	11	16	10	11	12			14,19		13,14	X		21,23
# DT2000012	10,13	30	11	10,12	15,17	11	10,11	12			17,18		16	X,Y		22,25
# DT2000007	11,15	30,32.2	11,12	11,12	15,17	7,10	8	11,13			17,19		15,16	X,Y		21,24
#########################################################################################################################################################################################################################################
        while (<IN>){
                chomp;
                my @str = split "\t", $_;
                ###### 这里需要测试，以确认处理逻辑 ##########
                unless (exists $exp_id{$str[0]}){  # 已有数据文件里的 实验编码 在 供患信息.txt 里不存在，则跳过该实验编码
                        # print "Next!\n" if $str[0] eq 'STR154465';
                        next;
                }
                ###很重要###
                ###### 这里需要测试，以确认处理逻辑 ##########
                if (exists $HasChimerism{$str[0]}){  # 判断 %HashChimerism (.PrevChimerism.txt) 中是否已有该 实验编码，如果有，则跳过该实验编码
                        print $HasChimerism{$str[0]},"\n";
                        next;
                }
                ###此处如果不写，会导致新的Allele无法读取###
                my $num = shift @str;  # 获取 实验编码，移除实验编码之后， @str 中仅存在所有marker的型别
                foreach my $tmp(@markers_jrk){  # 遍历所有marker  ** 注意，这样写，需要保证已有数据中的marker顺序 与 @markers_jrk中的顺序完全一致，否则会搞混。**
                        $PrevAllele{$num}{$tmp} = shift @str;  # 保存每个 实验编码 的每个marker 对应的 型别信息
                        # print $num,"|",$tmp,"|", $PrevAllele{$num}{$tmp} ,"\n" if $PrevAllele{$num}{$tmp}=~ /\s/;
                        $PrevAllele{$num}{$tmp} =~ s/\s//g;  # 去掉 型别信息 中的空字符
                }
        }
        close IN;
        # print $PrevAllele{STR154465}{D7S820},"\n";
        $error = "读取成功！";
        $text23 -> Text($error);  # 将 "已有数据" 部分列表框下方的 "尚未读取" 更新为 "读取成功！"
        $curr_index = 0;

        $SummaryLoaded = 1;  # 将 已有数据读取状态设置为 1，此时 $InputLoaded = 1 and $SummaryLoaded = 1
        $display2->Enable($InputLoaded*$SummaryLoaded);  # 将 "打印已有分型" 按钮，设置为"可点击"状态
        $open3->Enable($InputLoaded*$SummaryLoaded);  # 将 "添加下机数据" 按钮，设置为可点击状态
}

###################################################################################
# List2_MouseMove 函数: "已有数据" 部分的列表框鼠标移上时的处理函数                      #
###################################################################################
sub List2_MouseMove{
        $sb -> Text('双击条目以读取');
}

###################################################################################
# List2_MouseOut 函数: "已有数据" 部分的列表框鼠标移出时的处理函数                       #
###################################################################################
sub List2_MouseOut{
        $sb -> Text('');
}

###################################################################################
# Open2_Click 函数: "已有数据" 部分列表框下方的 "其他位置" 按钮点击时的处理函数            #
###################################################################################
sub Open2_Click{
        my @parms;
        push @parms,
          -filter =>
                [ 'TXT - Tab分隔文本', '*.txt'
                ],
          -directory => $ConfigHash{SummaryLoc},
          -title => '选择文件',
          -parent => $main,
          -owner => $main;
        my @file = Win32::GUI::GetOpenFileName ( @parms );
        return 0 unless $file[0];

        my $already = -1;
        foreach my $i(0 .. $#PrevFound){
                my $tmp = $PrevFound[$i];
                chomp $tmp;
                # print $tmp,"\n";
                # print $pwd.'\\'.$tmp,"\n";
                if ($tmp eq $file[0] or $file[0] eq $ConfigHash{SummaryLoc}.'\\'.$tmp){
                        $already = $i;
                        last;
                }
        }

        if ($already >= 0){
                $Input2 -> SetCurSel($already);
                return 0;
        }

        push @PrevFound, $file[0];
        my $tmp = &Shorten($file[0], 37);
        $file[0] =~ /(^.+)\\[^\\]+$/;
        $ConfigHash{SummaryLoc} = $1;

        $Input2 -> InsertString($tmp);
        push @PrevList, $tmp;

}

###################################################################################
# Open2_MouseMove 函数: "已有数据" 部分列表框下方的 "其他位置" 按钮鼠标移上时的处理函数  #
###################################################################################
sub Open2_MouseMove{
        $sb -> Text('从其他位置读取');
}

###################################################################################
# Open2_MouseMove 函数: "已有数据" 部分列表框下方的 "其他位置" 按钮鼠标移出时的处理函数   #
###################################################################################

sub Open2_MouseOut{
        $sb -> Text('');
}

###################################################################################
# DISPLAY2_Click 函数: "打印已有分型" 按钮 点击时的处理函数                            #
###################################################################################
### "打印已有分型" 指的是，打印已有 "术前患者" 和 "术前供者" 的型别 #######################
sub DISPLAY2_Click{
        unless (%PrevAllele){  # 已有分型结果为空，报错，提示用户先读取文件
                $error = "尚未读取文件！\n";
                Win32::MsgBox $error, 0, "错误！";
                return 0;
        }

A:
        my $Temp_typinglist = sprintf "分型列表-%4d%02d%02d.xlsx",$year, $mon, $mday;  # 定义 已有分型输出结果
        my $workbook;
        unless ($workbook = Excel::Writer::XLSX->new($Temp_typinglist)){  # 如果 excel文件 打不开，则报错，输出提示信息
                $error = $Temp_typinglist."正在使用中！请关闭后重试！";
                Win32::MsgBox $error, 0, "错误！";
                return 0;
        }

        my $format1 = $workbook->add_format(  # 定义 excel 表格的格式
                        size            => 9,
                        bold            => 0,
                        align           => 'left',
                        font            => decode('GB2312','宋体'),
                        'top'           => 1,
                        'bottom'        => 1,
                        'left'          => 1,
                        'right'         => 1,
        );

        my $worksheet = $workbook->add_worksheet();  # 往 excel文件中 添加 数据表(worksheet)
        $worksheet->hide_gridlines();  # 隐藏网格线
        $worksheet->keep_leading_zeros();  # 保留数字开头的 '0'
        $worksheet->set_landscape();  # 设置worksheet的页面方向（打印方向）为竖向
        $worksheet->set_paper(9);  # 设置打印纸的格式为 A4
        $worksheet->set_margin_left(0.394);  # 设置worksheet的左边距
        $worksheet->set_margin_right(0.394); # 设置worksheet的右边距
        $worksheet->set_column(0,2, 10);  # 设置0-2列 (第1-3列，共3列)的宽度为 10
        $worksheet->set_column(3,3, 1.75);  # 设置3列 (第4列)的宽度为 1.75
        $worksheet->set_column(4,5, 10);  # 设置4-5列 (第5-6列)的宽度为 10
        $worksheet->set_column(6,7, 1.75);  # 设置6-7列 (第7-8列)的宽度为 1.75
        $worksheet->set_column(7,8, 10);  # 设置7-8列 (第8-9列)的宽度为 10
        $worksheet->set_column(9,9, 1.75);  # 设置9列 (第10列)的宽度为 1.75
        $worksheet->set_column(10,11, 10);  # 设置10-11列 (第11-12列)的宽度为 10
        $worksheet->set_column(12,12, 1.75);  # 设置12列 (第13列)的宽度为 1.75
        $worksheet->set_column(13,15, 10);  # 设置13-15列 (第14-16列)的宽度为 10

        my $pages = int(($#TCA_id+1) / 10)+1;  # 每个页面打印10个报告单对应的 "术前患者"和"术前供者"的型别，@TCA_id:存储 供患信息.txt 中的报告单编号

        foreach my $i(0..$pages*39-1){  # 每个打印页面包括 40行
                $worksheet->set_row($i, 12.7);  # 设置 行高 为 12.7
        }

        foreach my $i(1..$pages){  # 为每一页的第一列和最后一列，写上 marker名
                # $worksheet->write(($i-1)*38,0,' ', $format1);
                # $worksheet->write(($i-1)*38+1,0,' ', $format1);
                # $worksheet->write(($i-1)*38+19,0,' ', $format1);
                # $worksheet->write(($i-1)*38+20,0,' ', $format1);
                my $j = 2;
                foreach (@markers_jrk){
                        $worksheet->write(($i-1)*38+$j,0,$markers_jrk[$j-2], $format1);
                        $worksheet->write(($i-1)*38+$j,15,$markers_jrk[$j-2], $format1);
                        $worksheet->write(($i-1)*38+$j+19,0,$markers_jrk[$j-2], $format1);
                        $worksheet->write(($i-1)*38+$j+19,15,$markers_jrk[$j-2], $format1);
                        $j ++;
                }
        }

        foreach my $i(0.. $#TCA_id){  # 遍历每个 报告单编号，输出对应的 "术前患者""术前供者" 的型别
                my $TCAID = $TCA_id[$i];  # 获取 报告单编号
                my @seq = split ",", $exp_seq{$TCAID};  # 获取该 报告单编号 对应的 "术前患者" "术前供者" ("术后患者" 可能包括也可能不包括)在 供患信息.txt 中的序号
                my $AAA = $exp_list[$seq[0]];  # 获取 "术前患者" 对应的 实验编码
                my $BBB = $exp_list[$seq[1]];  # 获取 "术前供者" 对应的 实验编码
                my $j = 2;
                my $strA;
                my $strB;
                # 每一行打印5个 报告单编号 对应的 "术前患者""术前供者" 的型别
                $worksheet->write(int($i/5)*19,$i%5*3+1,$data_in[$seq[-1]][7], $format1);  # 第0行 第1列，写上 "术前患者" 的实验编码
                $worksheet->write(int($i/5)*19,$i%5*3+2,decode('GB2312', $data_in[$seq[-1]][9]), $format1);  # 第0行 第2列，写上 "术前患者" 的姓名
                $worksheet->write(int($i/5)*19+1,$i%5*3+1,$AAA, $format1);  # 第1行 第1列，写上 "术前患者" 的型别
                $worksheet->write(int($i/5)*19+1,$i%5*3+2,$BBB, $format1);  # 第1行 第2列，写上 "术前供者" 的型别
                foreach (@markers_jrk){  # 遍历写入每个 marker 的型别
                        unless (exists $PrevAllele{$AAA}){  # "术前患者" 对应的 实验编码在 已有数据 中不存在  # %PrevAllele 保存已有数据中 每个 实验编码 的每个marker 对应的 型别信息
                                $strA = ' ';  # 不存在，则将 $strA = ' '
                        }else{  # "术前患者" 对应的 实验编码在 已有数据 中存在
                                $strA =  $PrevAllele{$AAA}{$_};  # 将 $strA 设置为 已有数据 中该marker对应的型别
                        }
                        unless (exists $PrevAllele{$BBB}){  # "术前供者" 对应的 实验编码在 已有数据 中不存在
                                $strB = ' ';  # 不存在，则将 $strB = ' '
                        }else{  # "术前供者" 对应的 实验编码在 已有数据 中存在
                                $strB =  $PrevAllele{$BBB}{$_};  # 将 $strB 设置为 已有数据 中该marker对应的型别
                        }
                        $worksheet->write(int($i/5)*19+$j,$i%5*3+1, decode('GB2312', $strA), $format1);  # 遍历写上 "术前患者" 的每个marker的型别
                        $worksheet->write(int($i/5)*19+$j,$i%5*3+2, decode('GB2312', $strB), $format1);  # 遍历写上 "术前患者" 的每个marker的型别
                        $j ++;
                }
        }

        $workbook -> close();  # 已有分型 写入 excel文件 "分型列表-%4d%02d%02d.xlsx" 完成
        `start $Temp_typinglist`;  # 调用windows cmd, 启动单独的“命令提示符”窗口来运行指定程序或命令。 实际效果：打开 "分型列表-%4d%02d%02d.xlsx"

        return 0;
}

###################################################################################
# DISPLAY2_Click 函数: "打印已有分型" 按钮 鼠标移上 时的处理函数                       #
###################################################################################
sub DISPLAY2_MouseMove{
        $sb -> Text('参考已有分型结果打印本次实验数据');
}

###################################################################################
# DISPLAY2_Click 函数: "打印已有分型" 按钮 鼠标移出 时的处理函数                       #
###################################################################################
sub DISPLAY2_MouseOut{
        $sb -> Text('');
}

#############################################################
################## 老版本功能，已弃用 ##########################
# sub W2_PREV_Click{
        # $direct = -1;
        # $curr_index --;
        # $dpwindow -> Hide();
        # $display2 -> Click();
        # return 0;
# }

# sub W2_NEXT_Click{
        # $direct = 1;
        # $curr_index ++;
        # $dpwindow -> Hide();
        # $display2 -> Click();
        # return 0;
# }

# sub W2_CLOSE_Click{
        # $dpwindow->Hide();
        # return 0;
# }

# sub W2_Terminate{
        # $dpwindow->Hide();
        # return 0;
# }

# sub W2_Resize {
    # my ($width, $height) = ($dpwindow->GetClientRect)[2..3];
    # $Grid->Resize ($width, $height-120);
# }

################## 老版本功能，已弃用 ##########################
#############################################################

###################################################################################
# Open3_Click 函数: "添加下机数据" 按钮 点击 时的处理函数                             #
###################################################################################
sub Open3_Click{
        unless (@exp_list){  # 添加下机数据 前需先读取供患信息
                $error = "请先读取供患信息！\n";
                Win32::MsgBox $error, 0, "错误！";
                return 0;
        }

        unless (%PrevAllele){  # 添加下机数据 前需先读取 已有分型信息
                $error = "请先读取已有分型信息！\n";
                Win32::MsgBox $error, 0, "错误！";
                return 0;
        }

        my @parms;
        push @parms,
          -multisel => 20,
          -filter =>
                [ 'TXT - Tab分隔文本', '*.txt'
                ],
          -directory => $ConfigHash{ThisLoc},  # 从 ThisLoc 目录下读取 .txt
          -title => '选择文件',
          -parent => $main,
          -owner => $main;
        my @file = Win32::GUI::GetOpenFileName ( @parms );  # 获取选中的文件名
        print "L1322:$_\n" for @file;
        return 0 unless $file[0];
        if (@file == 1){  # 选择一个文件
                chomp $file[0];
                push @ThisFound, $file[0];  # 将选中的 下机数据.txt 名字 存入 @ThisFound
                push @ThisList, &Shorten($file[0], 57);  # 调用 Shorten函数截短输入文件名，同时 存入 @ThisList
                $Input3 -> Enable(1);  # 将 "添加下机数据" 部分的列表框设置为可见
                $Input3 -> Add(&Shorten($file[0], 57));  # 将截断的文件名 显示在 "添加下机数据" 部分的列表框中
                $Read3->Enable(1);  # 将 "读取" 按钮 设置为 可点击
                return 0;
        }
        ## 以下处理 选中多个文件 的情况
        ##如果多选，返回格式为 路径;文件名1;文件名2...
        my $tmp = shift @file;
        chomp $tmp;
        $ConfigHash{ThisLoc} = $tmp;
        for my $i(0..$#file){
                chomp $file[$i];
                $file[$i] = $tmp."\\".$file[$i];
        }
        for (@file){
                my $already=-1;
                foreach my $i(0 .. $#ThisFound){
                        my $tmp = $ThisFound[$i];
                        chomp $tmp;
                        if ($_ eq $tmp){
                                $already = $i;
                                last;
                        }
                }

                if ($already >= 0){
                        $Input1 -> SetCurSel($already);
                        next;
                }

                push @ThisFound, $_;
                push @ThisList, &Shorten($_, 57);
                $Input3 -> Enable(1);
                $Input3 -> Add(&Shorten($_, 57));
        }

        $Read3->Enable(1);
}

###################################################################################
# Open3_MouseMove 函数: "添加下机数据" 按钮 鼠标移上 时的处理函数                      #
###################################################################################
sub Open3_MouseMove{
        $sb -> Text('选择下机数据文件并添加到右侧列表中');
}

###################################################################################
# Open3_MouseOut 函数: "添加下机数据" 按钮 鼠标移出 时的处理函数                       #
###################################################################################
sub Open3_MouseOut{
        $sb -> Text('');
}

###################################################################################
# List3_SelChange 函数: "添加下机数据" 部分的列表框 选中文件 时的处理函数               #
###################################################################################
sub List3_SelChange{
        my @sel = $Input3->GetSelItems();
        if (@sel > 0){
                $del3 -> Enable(1);  # 选中文件时，"移除" 按钮设置为 "可点击"
        }else{
                $del3 -> Enable(0);  # 未选中文件时，"移除" 按钮设置为 "不可点击"
        }
}

###################################################################################
# List3_MouseMove 函数: "添加下机数据" 部分的列表框 鼠标移上 时的处理函数               #
###################################################################################
sub List3_MouseMove{
        $sb -> Text('选定条目进行更多操作(支持Ctrl、Shift进行多选)');
}

###################################################################################
# List3_MouseOut 函数: "添加下机数据" 部分的列表框 鼠标移出 时的处理函数                #
###################################################################################
sub List3_MouseOut{
        $sb -> Text('');
}


################################################  2020.05.22.11.48 ################################################
###################################################################################
# Read3_Click 函数: "读取" 按钮 点击 时的处理函数                                    #
###################################################################################
# 下机数据文件 的格式如下 ##############################################################################################################################################################################################################
# 0Sample Name	1Panel	2Marker	3Allele 1	4Allele 2	5Allele 3	6Allele 4	7Peak Area 1	8Peak Area 2	9Peak Area 3	10Peak Area 4	11PHR	12AN
# 示例如下：
# Sample Name	Panel	Marker	Allele 1	Allele 2	Allele 3	Allele 4	Peak Area 1	Peak Area 2	Peak Area 3	Peak Area 4	PHR	AN
# 20FCM00134	Sinofiler_v1	D8S1179	10	11			41782	38113			0	0
# 20FCM00134	Sinofiler_v1	D21S11	29	30			19869	17794			0	0
# 20FCM00134	Sinofiler_v1	D7S820	12				41169				-2	0
# 20FCM00134	Sinofiler_v1	CSF1PO	9	13			30706	27814			0	0
# 20FCM00134	Sinofiler_v1	D3S1358	15	17			66461	67763			0	0
# 20FCM00134	Sinofiler_v1	D5S818	10	12	14		554	69030	64854		-1	-1
# 20FCM00134	Sinofiler_v1	D13S317	10	11			63975	57160			0	0
# 20FCM00134	Sinofiler_v1	D16S539	11	13			63035	55449			0	0
# 20FCM00134	Sinofiler_v1	D2S1338	17	19	24		307	45475	37783		-1	-1
# 20FCM00134	Sinofiler_v1	D19S433	12.2	15.2			32280	35424			0	0
# 20FCM00134	Sinofiler_v1	vWA	14	18			64017	57671			0	0
# 20FCM00134	Sinofiler_v1	D12S391	19	23			55455	45320			0	0
# 20FCM00134	Sinofiler_v1	D18S51	14				115992				-2	0
# 20FCM00134	Sinofiler_v1	AMEL	X				113366				-1	-1
# 20FCM00134	Sinofiler_v1	D6S1043	12	14			48624	44775			0	0
# 20FCM00134	Sinofiler_v1	FGA	18	23			38932	30702			0	0
# 20FCM00134-T	Sinofiler_v1	D8S1179	10	11	14		57390	45581	249		-1	-1
# 20FCM00134-T	Sinofiler_v1	D21S11	29	30			33034	26495			0	0
#########################################################################################################################################################################################################################################
sub Read3_Click{
        if (%ThisAllele){
                $error = "已经成功读取数据，是否重新读取？";
                my $s = Win32::MsgBox $error,1, "注意！";
                return 0 if $s != 1;
        }

        %ThisAllele = ();
        $ExpLoaded = 0;
        $run4 -> Enable(0);
        $text3 -> Text("尚未读取");

        foreach my $file (@ThisFound){  # 遍历 "添加下机数据"列表框中的文件
                next if $file =~/^\./;
                if (open IN,$file){

                }else{
                        $error = $file."打开失败！\n";
                        Win32::MsgBox $error, 0, "错误！";
                        return 0;
                };
                my %USE = ();
                while(<IN>){
                        chomp;
                        my $line = $_ ; print "L1396:$line\n";
                        # next if /Sample\sName/;  # 表头行 "Sample Name"开头
                        my $str_header = decode("utf8", "样本名称"); print "L1397:$str_header\n";
                        # next if /$str_header/;  # 表头行 "样本名称" 开头
                        if ($line =~ /$str_header/){
                            print "L1408:样本名称\n";
                            next;
                        }
                        next if $line =~ /LADDER/;  # 跳过 LADDER行 （内标）
                        next if $line =~ /NC\s/;  # 跳过含有 NC 的行，空对照
                        next if $line =~ /QC\d+\s/;  # 跳过含有 QC 的行，质控品? （未启用）

                        my @line = split /\t/,$line;
                        my ($tmpallele, $tmparea, $num);
                        if ($line[2] =~ /vWA/){$line[2] =~ s/vWA/VWA/;}  # 将marker "vWA" 替换为 "VMA"
                        if ($line[2] =~ /AMEL/){$line[2] =~ s/AMEL/Amel/;}  # 将marker "AMEL" 替换为 "Amel"
                        my $found = 0; print "L1404:$line[0]\n" ;
                        if (exists $exp_id{$line[0]}){  # $line[0]：在 供患信息.txt 中对应"实验编码" ，但在 下机数据.txt 中对应 "Sample Name"
                                $num = $line[0];  # $num：存放在 供患信息.txt 中存在的 "实验编码"
                                $found = 1;
                                # print "$num 找到了！\n";
                        }elsif($line[0] =~ /^(TB\d+)/){  # 判断 实验编码 是否以 TB+数字 格式开头
                                my $tmpstr = $1;  print "L1451:" . $tmpstr . "\n";  # 测试看看

                                if (exists $trans{$tmpstr}){  # %trans中存在 该实验编码的记录 #用来保存实验编码缩写到全称的转换
                                        if ($trans{$tmpstr} eq "ERROR"){
                                                $num = $line[0];
                                        }else{
                                                $num = $trans{$tmpstr};
                                                $found = 1;
                                        }
                                }else{  # %trans中 不存在 该实验编码的记录
                                        foreach my $str(@exp_list){  # @exp_list：用于存储 供患信息.txt 中的实验编码
                                                if ($str =~ /$tmpstr$/i){  # 判断 下机数据中读取的 实验编码 是否为 供患信息.txt 中某个实验编码的一部分 （即，简写）
                                                        $found = 1;
                                                        $num = $str;  # 将 实验编码简写 替换为 供患信息.txt 中的"实验编码"
                                                        $trans{$tmpstr} = $str;  # 记录 实验编码简写 与 供患信息.txt 中"实验编码" 的对应关系
                                                        # print "$tmpstr --> $str\n";
                                                        last;  # 跳出 遍历 @exp_list 的循环
                                                }
                                        }
                                        if ($found == 0){  # 实验编码 不是 供患信息.txt 中某个实验编码的一部分，提示 实验编码错误。
                                                # print "未找到",$line[0],"的实验记录！\n";
                                                $trans{$tmpstr} = "ERROR";  # 将该 实验编码 对应的 全称 设置为 "ERROR"
                                                $num = $line[0];  # 将实验编码 设置为 第一列的完整信息
                                        }
                                }
                        }elsif($line[0] =~ /^(\d{3,7})-?([A-Z]*)$/){  # 若实验编码的格式为 3-7位数字 + -(有/无) + 任意个数的大写字母
                                my $tmpstr = $2 ? $1.'-'.$2 : $1;  # 存在1个以上的大写字母，则将 $tmpstr 设置为 数字-字母；否则，将 $tmpstr 设置为 数字

                                if (exists $trans{$tmpstr}){  # %trans中存在 该实验编码的记录 #用来保存实验编码缩写到全称的转换
                                        if ($trans{$tmpstr} eq "ERROR"){
                                                $num = $line[0];
                                        }else{
                                                $num = $trans{$tmpstr};
                                                $found = 1;
                                        }
                                }else{  # %trans中 不存在 该实验编码的记录
                                        foreach my $str(@exp_list){  # @exp_list：用于存储 供患信息.txt 中的实验编码
                                                if ($str =~ /$tmpstr$/i){  # 判断 下机数据中读取的 实验编码 是否为 供患信息.txt 中某个实验编码的一部分 （即，简写）
                                                        $found = 1;
                                                        $num = $str;  # 将 实验编码简写 替换为 供患信息.txt 中的"实验编码"
                                                        $trans{$tmpstr} = $str;  # 记录 实验编码简写 与 供患信息.txt 中"实验编码" 的对应关系
                                                        # print "$tmpstr --> $str\n";
                                                        last;  # 跳出 遍历 @exp_list 的循环
                                                }
                                        }
                                        if ($found == 0){  # 实验编码 不是 供患信息.txt 中某个实验编码的一部分，提示 实验编码错误。
                                                # print "未找到",$line[0],"的实验记录！\n";
                                                $trans{$tmpstr} = "ERROR";  # 将该 实验编码 对应的 全称 设置为 "ERROR"
                                                $num = $line[0];  # 将实验编码 设置为 第一列的完整信息
                                        }
                                }
                        }else{  # 否则，输出提示信息，跳过该行
                                print "L1503:实验编码",$line[0],"有错误，请检查！\n";
                                next;
                        }

                        next if $found == 0;  # 实验编码 不是 供患信息.txt 中某个实验编码(或它的的一部分)，提示 实验编码错误，跳过该行。
                        #print $file,"|",$num,"\n";
# 下机数据文件 的格式如下 ##############################################################################################################################################################################################################
# 0Sample Name	1Panel	2Marker	3Allele 1	4Allele 2	5Allele 3	6Allele 4	7Peak Area 1	8Peak Area 2	9Peak Area 3	10Peak Area 4	11PHR	12AN
# 示例如下：
# Sample Name	Panel	Marker	Allele 1	Allele 2	Allele 3	Allele 4	Peak Area 1	Peak Area 2	Peak Area 3	Peak Area 4	PHR	AN
# 20FCM00134	Sinofiler_v1	D8S1179	10	11			41782	38113			0	0
                        if    ($line[6]){$tmpallele = join ",", ($line[3],$line[4],$line[5],$line[6]);}  # $line[6] = Allele 4 不为空
                        elsif ($line[5]){$tmpallele = join ",", ($line[3],$line[4],$line[5]);}  # Allele 4 为空 && Allele 3 不为空
                        elsif ($line[4]){$tmpallele = join ",", ($line[3],$line[4]);}  # Allele 4 和 Allele 3 都为空 && Allele 2  不为空
                        else            {$tmpallele =             $line[3];}  # Allele 4 , Allele 3 和 Allele 2 都为空

                        if    ($line[10]){$tmparea = join ",", ($line[7],$line[8],$line[9],$line[10]);}  # Peak Area 4 不为空
                        elsif ($line[9]) {$tmparea = join ",", ($line[7],$line[8],$line[9]);}  # Peak Area 4 为空 && Peak Area 3 不为空
                        elsif ($line[8]) {$tmparea = join ",", ($line[7],$line[8]);}  # Peak Area 4 和 Peak Area 3 都为空 && Peak Area 2 不为空
                        else             {$tmparea =             $line[7];}  # Peak Area 4, 3,2 都为空

                        # if (exists $PrevAllele{$num}{$line[2]}){
                                # $ThisAllele{$num}{$line[2]} = $tmpallele;
                                # $area  {$num}{$line[2]} = $tmparea;
                        # }
                        #
                        if (exists $PrevAllele{$num}){  # %PrevAllele 保存已有数据中 每个 实验编码 的每个marker 对应的 型别信息
                                # 实验编码 在 已有数据中已经存在型别信息
                                print "$num 已有！\n";
                                if (exists $USE{$num}){
                                        next;
                                }else{
                                        $error = "实验编码:".$num." 已有分型数据，本次下机数据是否使用？如果此样本是患者术后，本次数据不使用将会导致错误！";
                                        my $s = Win32::MsgBox $error,4, "注意！";
                                        if ($s == 6){  # 用户选择 使用本次下机数据 $s == 6 means user select "Yes"
                                                delete $PrevAllele{$num};  # 删除 已有分析数据 中该实验编码的记录
                                                $ThisAllele{$num}{$line[2]} = $tmpallele;  # 存放 该实验编码 该marker 对应的型别信息
                                                $area{$num}{$line[2]} = $tmparea;  # 存放 该实验编码 该marker 对应的 area信息
                                        }else{  # 用户选择 不适用本次下机数据
                                                $USE{$num} = 'no';
                                                next;  # 跳到下一行
                                        }
                                }
                        }else{  # 实验编码 在 已有数据 中 不存在型别信息
                                $ThisAllele{$num}{$line[2]} = $tmpallele;  # 存放 该实验编码 该marker 对应的型别信息
                                $area{$num}{$line[2]} = $tmparea;  # 存放 该实验编码 该marker 对应的 area信息
                        }
                        #


                        #print $file,"|",$num,"|",$line[2],"|",$allele{$num}{$line[2]},"|",$area{$num}{$line[2]},"\n";

                }
                close IN;

        }
        $error = "读取成功";
        $text3 -> Text($error);  # 将 "添加下机数据" 列表框下方"尚未读取" 提示信息更新为 "读取成功"
        $ExpLoaded = 1;  # 标记 下机数据已经读取
        $run4 -> Enable(1);  # 将 "生成报告" 按钮设置为 "可点击"
        return 0;
}

###################################################################################
# Read3_MouseMove 函数: "读取" 按钮 鼠标移上 时的处理函数                             #
###################################################################################
sub Read3_MouseMove{
        $sb -> Text('读取左侧列表中的数据');
}

###################################################################################
# Read3_MouseOut 函数: "读取" 按钮 鼠标移出 时的处理函数                              #
###################################################################################
sub Read3_MouseOut{
        $sb -> Text('');
}

###################################################################################
# Del3_Click 函数: "移除" 按钮 点击 时的处理函数                                      #
###################################################################################
sub Del3_Click{
        my @sel = $Input3->GetSelItems();
        # print $_,"," for @sel;
        # print "\n";
        my $index;
        $Input3 -> DeleteString($sel[0]);
        @ThisFound = &DelItem(@ThisFound, $sel[0]);
        @ThisList = &DelItem(@ThisList, $sel[0]);
        foreach my $i(1..$#sel){
                $index = $sel[$i] - $i;
                $Input3 -> DeleteString($index);
                @ThisFound = &DelItem(@ThisFound, $index);
                @ThisList = &DelItem(@ThisList, $index);
        }
        if (@ThisFound == 0){
                %ThisAllele = ();
                $ExpLoaded = 0;
                $run4 -> Enable(0);
                $text3 -> Text("尚未读取");
                $Read3 -> Enable(0);
                $Input3 -> Enable(0);
                $del3 -> Enable(0);
        }
}

###################################################################################
# Del3_MouseMove 函数: "移除" 按钮 鼠标移上 时的处理函数                              #
###################################################################################
sub Del3_MouseMove{
        $sb -> Text('移除右侧列表中选中的文件');
}

###################################################################################
# Del3_MouseOut 函数: "移除" 按钮 鼠标移出 时的处理函数                              #
###################################################################################
sub Del3_MouseOut{
        $sb -> Text('');
}

###################################################################################
# RUN_Click 函数: "生成报告" 按钮 点击 时的处理函数                                    #
###################################################################################
sub RUN_Click{
        my $ret = Win32::GUI::BrowseForFolder (  # 选择 报告存放 的目录
                -title      => "请选择保存路径",
                # -editbox    => 1,
                -directory  => $ConfigHash{OutputLoc},
                -folderonly => 1,
                -newui      => 1,
                -parent => $main,
                -owner => $main,
        );
        return 0 unless $ret;
        $Output_Dir = $ret;
        $ConfigHash{OutputLoc} = $ret;  # 将 $ConfigHash{OutputLoc} 更新为 $ret

        $sb->Move( 0, ($main->ScaleHeight() - $sb->Height()) );
        $sb->Resize( $main->ScaleWidth(), $sb->Height() );
        $sb->Text("正在合并处理文件...");
        $RUNwindow -> Show();
        #################### 定义一组变量用于存储 汇总报告单中需要的信息 ########################################################################################
        my %this_patient_ID_and_report_id = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者本轮检测的汇总报告单的编号
        my %this_patient_ID_and_patient_name = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者姓名
        my %this_patient_ID_and_patient_gender = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者性别
        my %this_patient_ID_and_patient_age = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者年龄
        my %this_patient_ID_and_patient_diagnosis = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者 诊断信息
        my %this_patient_ID_and_patient_sampleType = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者样本类型 (外周血 / 骨髓 /骨髓血?)
        # my %this_patient_ID_and_patient_sampleDetailType = () ;  # 定义一个hash表，用户存储 患者编码 (如:HUN001胡琳) 对应的 患者样本详细类型  (骨髓血-B细胞分选...)
        my %this_patient_ID_and_patient_sampleDate = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者采样日期
        my %this_patient_ID_and_patient_rcvDate = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者收样(接样)日期
        my %this_patient_ID_and_donor_name = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 供者姓名
        my %this_patient_ID_and_donor_gender = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 供者性别
        my %this_patient_ID_and_donor_age = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 供者年龄
        my %this_patient_ID_and_donor_relationship = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 供者与他/她的关系
        my %this_patient_ID_and_hospital = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 医院全称
        my %this_patient_ID_and_doctor = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 送样医生
        my %this_patient_ID_sampleDetailType_and_conclusion = () ;  # 定义一个hash表，用于存储 患者编码 某个样本详细类型 对应的 结论
        my %this_patient_ID_and_shuqian_donor_expid = () ;  # 定义一个hash表，用于存储 患者编码 对用的 术前供者 实验编码
        my %this_patient_ID_and_shuqian_patient_expid = () ;  # 定义一个hash表，用于存储 患者编码 对用的 术前患者 实验编码
        my %this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_expid = () ;  # 定义一个hash表，用于存储 患者编码 对用的 术后患者(外周血/骨髓血) 实验编码
        my %this_patient_ID_and_shuhou_patient_T_cell_expid = () ;  # 定义一个hash表，用于存储 患者编码 对应的 术后患者(T细胞分选) 实验编码
        my %this_patient_ID_and_shuhou_patient_B_cell_expid = () ;  # 定义一个hash表，用于存储 患者编码 对应的 术后患者(B细胞分选) 实验编码
        my %this_patient_ID_and_shuhou_patient_NK_cell_expid = () ;  # 定义一个hash表，用于存储 患者编码 对应的 术后患者(NK细胞分选) 实验编码
        my %this_patient_ID_and_shuhou_patient_li_cell_expid = () ;  # 定义一个hash表，用于存储 患者编码 对应的 术后患者(粒细胞分选) 实验编码
        my %hash_this_patient_ID_and_shuqian_patient_genotypes = () ;  # 定义一个hash表，用于存储 患者编码 对应的 术前患者 的 每个marker名 对应的 型别结果
        my %hash_this_patient_ID_and_shuqian_donor_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术前供者 的 每个marker名 对应的 型别结果
        my %hash_this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术后(外周血/骨髓血)  的 每个marker名 对应的 型别结果
        my %hash_this_patient_ID_and_shuhou_patient_T_cell_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术后(T细胞) 的 每个marker名 对应的 型别结果
        my %hash_this_patient_ID_and_shuhou_patient_B_cell_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术后(B细胞) 的 每个marker名 对应的 型别结果
        my %hash_this_patient_ID_and_shuhou_patient_NK_cell_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术后(NK细胞) 的 每个marker名 对应的 型别结果
        my %hash_this_patient_ID_and_shuhou_patient_li_cell_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术后(粒细胞) 的 每个marker名 对应的 型别结果

        %allele = ();

        foreach my $PrevKey1(keys %PrevAllele){  # %PrevAllele 保存已有数据中 每个 实验编码 的每个marker 对应的 型别信息
            # $PrevKey1: 实验编码
                foreach my $PrevKey2(keys %{$PrevAllele{$PrevKey1}}){  # $PrevKey2：marker
                        $allele{$PrevKey1}{$PrevKey2} = $PrevAllele{$PrevKey1}{$PrevKey2};  # 将已有数据中 每个实验编码 的每个marker 对应的型别信息，存入 %allele
                        # print "Prev $PrevKey1|$PrevKey2|",$allele{$PrevKey1}{$PrevKey2},"\n";
                }
        }
        foreach my $ThisKey1(keys %ThisAllele){  # %ThisAllele 存放 该实验编码 该marker 对应的型别信息
            # $ThisKey1: 下机数据文件中的 实验编码
                foreach my $ThisKey2(keys %{$ThisAllele{$ThisKey1}}){  # $ThisKey2：marker
                        $allele{$ThisKey1}{$ThisKey2} = $ThisAllele{$ThisKey1}{$ThisKey2};  # 将 下机数据中 每个实验编码 的每个marker 对应的型别信息，存入 %allele
                        # print "This $ThisKey1|$ThisKey2|",$allele{$ThisKey1}{$ThisKey2},"\n";
                }
        }

        my (%date4,%date1,%date2,%sample,%operation,%cells,%date3,%number,%rptnum,%name,%patient,%gender,%age,%diagnosis,%relation,%xnum,%hospital,%doctor,%hosptl_num,%bed_num);
        my %sheet_name;

        foreach (keys %exp_id){  # %exp_id，存放每个实验编码的原始顺序
                # 供患关系文件 的格式如下 ##############################################################################################################################################################################################################
                # 0收样日期	1生产时间	2移植日期	3样品类型	4样品性质	5分选类别	6采样日期	7实验编码	8报告单编号	9姓名	10供患关系	11性别	12年龄	13诊断	14亲缘关系	15关联样本编号	16医院编码	17送检医院	18送检医生  19住院号   20床号
                # 示例如下：
                # 收样日期	生产时间	移植日期	样品类型	样品性质	分选类别	采样日期	实验编码	报告单编号	姓名	供患关系	性别	年龄	诊断	亲缘关系	关联样本编号	医院编码	送检医院	送检医生  住院号   床号
                # 				术前			D19STR00039	QC-Q019	Q17	患者
                # 				术前			10751	QC-Q019		供者
                # 				术后			Q19	QC-Q019	Q19	患者							南京市儿童医院
                # 	2020/3/5		[其他]	术前		2020/3/3	D20STR01231	TCA2007498	吴久芳	患者	男	47	-	吴久芳	本人
                # 	2020/3/5		[其他]	术前		2020/3/3	D20STR01232	TCA2007498	吴文方	供者	男	44	-	弟弟
                # 2020/3/4	2020/3/5		骨髓血	术后		2020/3/3	D20STR01230	TCA2007498	吴久芳	患者	男	47	-	吴久芳	本人		广东省人民医院	黄励思
                # 	2018/6/8		全血	术前			STR1808282	TCA2007647	杨梅月	患者	女	不详		杨梅月	本人
                # 	2018/6/8		全血	术前			STR1810793	TCA2007647	杨梅	供者	女	不详		姐姐

                # 获取当前 实验编码 在供患信息.txt 里的原始顺序
                my $number = $exp_id{$_};  # $_ : 实验编码;

                $date4{$_}     = $data_in[$number][0];       # 供患信息中 实验编码 对应的 收样日期
                $date1{$_}     = $data_in[$number][1];       # 供患信息中 实验编码 对应的 生产时间
                $date2{$_}     = $data_in[$number][2];       # 供患信息中 实验编码 对应的 移植日期
                $sample{$_}    = $data_in[$number][3];       # 供患信息中 实验编码 对应的 样本类型
                $operation{$_} = $data_in[$number][4];       # 供患信息中 实验编码 对应的 样品性质
                $cells{$_}     = $data_in[$number][5];       # 供患信息中 实验编码 对应的 分选类型
                $date3{$_}     = $data_in[$number][6];       # 供患信息中 实验编码 对应的 采样日期
                $number{$_}    = $data_in[$number][7];       # 供患信息中 实验编码 对应的 实验编码
                $rptnum{$_}    = $data_in[$number][8];       # 供患信息中 实验编码 对应的 报告单编号
                $name{$_}      = $data_in[$number][9];       # 供患信息中 实验编码 对应的 姓名
                $patient{$_}   = $data_in[$number][10];      # 供患信息中 实验编码 对应的 供患关系
                $gender{$_}    = $data_in[$number][11];      # 供患信息中 实验编码 对应的 型别
                $age{$_}       = $data_in[$number][12];      # 供患信息中 实验编码 对应的 年龄
                $diagnosis{$_} = $data_in[$number][13];      # 供患信息中 实验编码 对应的 诊断
                $relation{$_}  = $data_in[$number][14];      # 供患信息中 实验编码 对应的 亲缘关系
                $xnum{$_}      = $data_in[$number][15];      # 供患信息中 实验编码 对应的 关联样本编号
                $hospital{$_}  = $data_in[$number][17];      # 供患信息中 实验编码 对应的 送检医院
                $doctor{$_}    = $data_in[$number][18];      # 供患信息中 实验编码 对应的 送检医生
                $hosptl_num{$_}= $data_in[$number][19];      # 供患信息中 实验编码 对应的 住院号
                $bed_num{$_}   = $data_in[$number][20];      # 供患信息中 实验编码 对应的 床号
        }

        $date4{'  '}     = '';
        $date1{'  '}     = '';
        $date2{'  '}     = '';
        $sample{'  '}    = '';
        $operation{'  '} = '';
        $cells{'  '}     = '';
        $date3{'  '}     = '';
        $number{'  '}    = '';
        $rptnum{'  '}    = '';
        $name{'  '}      = '';
        $patient{'  '}   = '';
        $gender{'  '}    = '';
        $age{'  '}       = '';
        $diagnosis{'  '} = '';
        $relation{'  '}  = '';
        $xnum{'  '}      = '';
        $hospital{'  '}  = '';
        $doctor{'  '}    = '';
        $hosptl_num{'  '}= '';
        $bed_num{'  '}   = '';
        foreach (@markers_jrk){$allele{'  '}{$_} = ' '; }
        foreach (@markers_jrk){$area  {'  '}{$_} = ' '; }


        $sb->Move( 0, ($main->ScaleHeight() - $sb->Height()) );
        $sb->Resize( $main->ScaleWidth(), $sb->Height() );
        $sb->Text("正在生成文件...");  # 状态条，输出提示信息

        my $success = 1;
        my @conclusion;  # 存放 第几份 报告的结论
        my @num1;  # 存放 第几份 报告对应的 "术前患者" 的实验编码
        my @num2;  # 存放 第几份 报告对应的 "术前供者" 的实验编码
        my @num3;  # 存放 第几份 报告对应的 "术后患者" 的实验编码
        my @sheet;  # 存放 第几份 报告对应的 "患者姓名"

        # my @count_sum;
        my @count_n;
        my @count_avg;
        my @SD;
        my @marker_type;
        my @type;
        my @count;
        # 遍历每个 报告单编号
        # $z: 第几个报告单编号
        foreach my $z(0 .. $#TCA_id){  # @TCA_id:存储 供患信息.txt 中的报告单编号
                my $TCAID = $TCA_id[$z];  # 获取 报告单编号
                if ($exp_error{$TCAID} == 1){  # 跳过 存在实验问题的报告单编号
                        $conclusion[$z] = '跳过';
                        next;
                }
                print STDERR "L1775:",$TCAID,"实验数：",$exp_num{$TCAID},"\n";
                $RptBox -> Append('准备'.$TCAID.'...');  # "生成报告" 部分的文本框中显示提示信息
                if ($exp_num{$TCAID} == 1){  # 该 "报告单编号" 对应的 实验编码 个数 为 1（即 该报告单对应的实验个数）
                        # 如果仅有一个 术后患者 样本，则不出报告，报错，输出提示信息
                        $conclusion[$z] = '无';
                        my @seq = split ",", $exp_seq{$TCAID};  # 获取该 报告单编号 对应的 "术前患者" "术前供者" ("术后患者" 可能包括也可能不包括)在 供患信息.txt 中的序号
                        unless (exists $allele{$exp_list[$seq[0]]}){  # 实验编码，在 %allele 中不存在
                                # %allele: 已有数据 和 下机数据 中，每个实验编码 的每个 marker 对应的 型别信息
                                # @exp_list：用于存储 供患信息.txt 中的实验编码
                                # $seq[0]:该 报告单编号 对应的 "术前患者" 在 供患信息.txt 中的序号
                                # $exp_list[$seq[0]]：获取该 报告单编号 对应的 "术前患者" 在 供患信息.txt 中的 实验编码
                                # $allele{$exp_list[$seq[0]]}：获取 "术前患者" 所有marker 对应的 型别信息，返回值为 数组

                                $error = '下机数据中未找到编号'.$exp_list[$seq[0]].'的数据请检查！';
                                Win32::MsgBox $error, 0, "错误！";
                                $success = 0;
                                $RptBox -> Append("失败【下机数据不全】\r\n");
                                $conclusion[$z] = '跳过';
                                next;
                        }
                        if ($data_in[$seq[0]][4] eq "术前"){  # 术前样本
                                if ($data_in[$seq[0]][10] eq "患者"){  # 术前患者
                                        $num1[$z] = $exp_list[$seq[0]];  # $exp_list[$seq[0]]：获取该 报告单编号 对应的 "术前患者" 在 供患信息.txt 中的 实验编码
                                        $num2[$z] = '  ';
                                        $num3[$z] = '  ';
                                        $sheet[$z] = $name{$num1[$z]};
                                }else{  # 术前供者
                                        $num1[$z] = '  ';
                                        $num2[$z] = $exp_list[$seq[0]];  # $exp_list[$seq[0]]：获取该 报告单编号 对应的 "术前患者" 在 供患信息.txt 中的 实验编码
                                        $num3[$z] = '  ';
                                        $sheet[$z] = $name{$num2[$z]};
                                }
                        }else{  # 术后样本
                                if ($data_in[$seq[0]][10] eq "患者"){  # 术后患者
                                        $num1[$z] = '  ';
                                        $num2[$z] = '  ';
                                        $num3[$z] = $exp_list[$seq[0]];
                                        $sheet[$z] = $name{$num3[$z]};
                                }else{  # 术后供者
                                        my $error  = "报告编号：".$TCAID."只包含一份显示为供者术后的样本。\n请检查，本次将不生成报告！\n";
                                        Win32::MsgBox $error, 0, "注意！";
                                        $RptBox -> Append("失败【术后供者】\r\n");
                                        $conclusion[$z] = '跳过';
                                        next;
                                }
                        }
                        $RptBox -> Append("成功！\r\n");
                        # printf "%s|%s|%s|%s|%s\n",        $TCAID, $num1[$z], $num2[$z], $num3[$z], $sheet[$z];
                }elsif ($exp_num{$TCAID} == 2){  # 报告单编号 对应 2个实验编码
                        $conclusion[$z] = '无';
                        my @seq = split ",", $exp_seq{$TCAID};  # 获取该 报告单编号 对应的 "术前患者" "术前供者" ("术后患者" 可能包括也可能不包括)在 供患信息.txt 中的序号
                        unless (exists $allele{$exp_list[$seq[0]]}){  # 第一个实验编码，在 %allele 中不存在
                                # %allele: 已有数据 和 下机数据 中，每个实验编码 的每个 marker 对应的 型别信息
                                # @exp_list：用于存储 供患信息.txt 中的实验编码
                                # $seq[0]:该 报告单编号 对应的 "术前患者" 在 供患信息.txt 中的序号
                                # $exp_list[$seq[0]]：获取该 报告单编号 对应的 "术前患者" 在 供患信息.txt 中的 实验编码
                                # $allele{$exp_list[$seq[0]]}：获取 "术前患者" 所有marker 对应的 型别信息，返回值为 数组

                                $error = '下机数据中未找到编号'.$exp_list[$seq[0]].'的数据请检查！';
                                Win32::MsgBox $error, 0, "错误！";
                                $success = 0;
                                $RptBox -> Append("失败【下机数据不全】\r\n");
                                $conclusion[$z] = '跳过';
                                next;
                        }
                        unless (exists $allele{$exp_list[$seq[1]]}){  # 第二个实验编码，在 %allele 中不存在
                                $error = '下机数据中未找到编号'.$exp_list[$seq[1]].'的数据请检查！';
                                Win32::MsgBox $error, 0, "错误！";
                                $success = 0;
                                $RptBox -> Append("失败【下机数据不全】\r\n");
                                $conclusion[$z] = '跳过';
                                next;
                        }

                        my $sum = 0;  # "术前患者" = 4 ; "术前供者" = 8; "术后患者" = 6; "术后供者" = 10
                        # "术前患者" + "术前供者" = 12
                        # "术前患者" + "术后患者" = 10
                        # "术前供者" + "术后患者" = "术前患者" + "术后供者" = 14
                        # "术后患者" + "术后供者" = 16
                        # "术前供者" + "术后供者" = 18
                        foreach my $i(@seq){
                                if ($data_in[$i][4] eq "术前"){
                                        $sum += 0;
                                }else{
                                        $sum += 2;
                                }
                                if ($data_in[$i][10] eq "患者"){
                                        $sum += 4;
                                }else{
                                        $sum += 8;
                                }
                        }
                        if ($sum == 12){  # "术前患者" + "术前供者"
                                $num1[$z] = $exp_list[$seq[0]];  # $exp_list[$seq[0]]：获取该 报告单编号 对应的 "术前患者" 在 供患信息.txt 中的 实验编码
                                $num2[$z] = $exp_list[$seq[1]];  # $exp_list[$seq[1]]：获取该 报告单编号 对应的 "术前供者" 在 供患信息.txt 中的 实验编码
                                $num3[$z] = '  ';
                                $sheet[$z] = $name{$num1[$z]};   # 存放 第几份 报告对应的 "患者姓名"
                        }elsif($sum == 10){  # "术前患者" + "术后患者"
                                $num1[$z] = $exp_list[$seq[0]];  # 存放 第几份 报告对应的 "术前患者" 在 供患信息.txt 中的 实验编码
                                $num2[$z] = '  ';
                                $num3[$z] = $exp_list[$seq[1]];  # 存放 第几份 报告对应的 "术后患者" 在 供患信息.txt 中的 实验编码
                                $sheet[$z] = $name{$num1[$z]};   # 存放 第几份 报告对应的 "患者姓名"
                        }elsif($sum == 14){  # "术前供者" + "术后患者" = "术前患者" + "术后供者" = 14  ("术前患者" + "术后供者" ??)
                                $num1[$z] = '  ';
                                $num2[$z] = $exp_list[$seq[0]];  # 存放 第几份 报告对应的 "术前供者" 在 供患信息.txt 中的 实验编码
                                $num3[$z] = $exp_list[$seq[1]];  # 存放 第几份 报告对应的 "术后患者" 在 供患信息.txt 中的 实验编码
                                $sheet[$z] = $name{$num3[$z]};   # 存放 第几份 报告对应的 "患者姓名"
                        }else{
                                my $error  = "报告编号：".$TCAID."包含一份显示为供者术后的样本。请检查，本次将不生成报告！";
                                Win32::MsgBox $error, 0, "注意！";
                                $RptBox -> Append("失败【术后供者】\r\n");
                                $conclusion[$z] = '跳过';
                                next;
                        }
                        $RptBox -> Append("成功！\r\n");
                        ####
                        #上面通过计数器来判断2个样本的情况分属患者还是供者
                        ####
                        # printf "%s|%s|%s|%s|%s\n",        $TCAID, $num1[$z], $num2[$z], $num3[$z],$sheet[$z];
                }elsif ($exp_num{$TCAID} == 3){  # 报告单编号 对应 3个实验编码
                        my @seq = split ",", $exp_seq{$TCAID};
                        print "L1846:$TCAID|",$exp_list[$seq[0]],"|", $allele{$exp_list[$seq[0]]}{D7S820},"\n";
                        print "L1847:$TCAID|",$exp_list[$seq[1]],"|", $allele{$exp_list[$seq[1]]}{D7S820},"\n";
                        print "L1848:$TCAID|",$exp_list[$seq[2]],"|", $allele{$exp_list[$seq[2]]}{D7S820},"\n";
                        print "L1849:$TCAID|", $conclusion[$z] , "\n";

                        # 检查 3个实验编码 在 %allele 中是否存在
                        unless (exists $allele{$exp_list[$seq[0]]}){
                                $error = '下机数据中未找到编号'.$exp_list[$seq[0]].'的数据请检查！';
                                Win32::MsgBox $error, 0, "错误！";
                                $success = 0;
                                $RptBox -> Append("失败【下机数据不全】\r\n");
                                $conclusion[$z] = '跳过';
                                next;
                        }
                        unless (exists $allele{$exp_list[$seq[1]]}){
                                $error = '下机数据中未找到编号'.$exp_list[$seq[1]].'的数据请检查！';
                                Win32::MsgBox $error, 0, "错误！";
                                $success = 0;
                                $RptBox -> Append("失败【下机数据不全】\r\n");
                                $conclusion[$z] = '跳过';
                                next;
                        }
                        unless (exists $allele{$exp_list[$seq[2]]}){
                                $error = '下机数据中未找到编号'.$exp_list[$seq[2]].'的数据请检查！';
                                Win32::MsgBox $error, 0, "错误！";
                                $success = 0;
                                $RptBox -> Append("失败【下机数据不全】\r\n");
                                $conclusion[$z] = '跳过';
                                next;
                        }

                        $num1[$z] = $exp_list[$seq[0]];  # 存放 第几份 报告对应的 "术前患者" 在 供患信息.txt 中的 实验编码
                        $num2[$z] = $exp_list[$seq[1]];  # 存放 第几份 报告对应的 "术前供者" 在 供患信息.txt 中的 实验编码
                        $num3[$z] = $exp_list[$seq[2]];  # 存放 第几份 报告对应的 "术后患者" 在 供患信息.txt 中的 实验编码
                        $sheet[$z] = $name{$num1[$z]};   # 存放 第几份 报告对应的 "患者姓名"
                        # printf "%s|%s|%s|%s|%s\n",        $TCAID, $num1, $num2, $num3,$sheet;
                }else{  # 超过 3个 实验编码的 报告单编号，报错，提示错误信息，同时跳过该 报告单
                        my $error  = "报告编号：".$TCAID."的实验样本编号过多！请检查！\n";
                        Win32::MsgBox $error, 0, "注意！";
                        $RptBox -> Append("失败【样本过多】\r\n");
                        $conclusion[$z] = '跳过';
                        next;
                }
                # print $num2,"|",$relation{$num2},"\n";
                # $count_sum[$z] = 0;
                $count_n[$z] = 0;  # 记录每份报告 有嵌合率结果的位点（有效位点）个数
                $count_avg[$z] = 0;  # 记录每份报告 有效位点的嵌合率 的平均值
                $SD[$z] = 0;  # 记录每份报告 有效位点的嵌合率 的SD

                my $errorcount = 0;  # 存在错误的marker位点个数
                # 遍历该 报告单编号 对应的每个marker位点
                foreach my $k (0..$#markers_jrk){

                        # if ($conclusion[$z]){
                                # if ($conclusion[$z] eq '跳过'){
                                        # last;
                                # }
                        # }
                        my %alleles_before = ();

                        # print $num1[$z],"|",$allele{$num1[$z]}{$markers[$k]},"\n";
                        # print $num2[$z],"|",$allele{$num2[$z]}{$markers[$k]},"\n";
                        # print $num3[$z],"|",$allele{$num3[$z]}{$markers[$k]},"\n";
                        # print $num3[$z],"|",$area{$num3[$z]}{$markers[$k]},"\n";

                        my @allele1 = split/,/, $allele{$num1[$z]}{$markers_jrk[$k]};  print "L1929:$z\t$k\t$allele{$num1[$z]}{$markers_jrk[$k]}\n"; # 获取该报告单 "术前患者" 每个位点的型别
                        $alleles_before{$_} = 1 foreach @allele1;  # 将 "术前患者" 每个位点出现的genotype 标记为 1. 用 %alleles_before 记录。
                        my @allele2 = split/,/, $allele{$num2[$z]}{$markers_jrk[$k]};  print "L1931:$z\t$k\t$allele{$num2[$z]}{$markers_jrk[$k]}\n"; # 获取该报告单 "术前供者" 每个位点的型别
                        $alleles_before{$_} = 1 foreach @allele2;  # 将 "术前供者" 每个位点出现的genotype 标记为 1. 用 %alleles_before 记录。
                        my @allele3 = split/,/, $allele{$num3[$z]}{$markers_jrk[$k]};  print "L1933:$z\t$k\t$allele{$num3[$z]}{$markers_jrk[$k]}\n"; # 获取该报告单 "术后患者" 每个位点的型别

                        # 判断 "术后患者" 是否存在 "术前患者" and "术前供者" 中均未出现的 genotype
                        foreach (@allele3){
                                if (!exists $alleles_before{$_}){  # 如果出现"术前患者" and "术前供者" 中均未出现的 genotype，则计数 存在错误的marker位点个数
                                        $type[$z][$k] = "error";  # 将 第几份报告单 的 第几个marker 位点的type状态 标记为 "error"
                                        $count[$z][$k] = "error";  # 将 第几份报告单 的 第几个marker 位点的count状态 标记为 "error"
                                        $errorcount ++;
                                        last;
                                }
                        }

                        my @area3   = split/,/, $area{$num3[$z]}{$markers_jrk[$k]};
                        # print $_,"|" foreach @allele1;
                        # print $_,"|" foreach @allele2;
                        # print $_,"|" foreach @allele3;
                        # print $_,"|" foreach @area3;
                        # print "\n";
                }

                # 如果 存在错误的位点个数 >= 10，则报错，提示错误信息，跳过该报告单
                if ($errorcount >= 10){  # Test, change 6 to 10
                        $success = 0;
                        $error = '报告单号'.$TCAID.'的25个位点中'.$errorcount.'个分型错误！请检查！将跳过出具此份报告。';
                        Win32::MsgBox $error, 0, "注意";
                        $RptBox -> Append("失败【分型数据错误】\r\n");
                        $conclusion[$z] = '跳过';
                        next;
                }

                # 遍历每个位点，判断每个位点的组成类型 存入 %type
                # 遍历每个位点，根据位点组成类型，计算其 嵌合率 存入 %count
                foreach my $k (0..$#markers_jrk){
                        next if $count[$z][$k] eq 'error';
                        my @allele1 = split/,/, $allele{$num1[$z]}{$markers_jrk[$k]};
                        my @allele2 = split/,/, $allele{$num2[$z]}{$markers_jrk[$k]};
                        my @allele3 = split/,/, $allele{$num3[$z]}{$markers_jrk[$k]};
                        my @area3   = split/,/, $area{$num3[$z]}{$markers_jrk[$k]};

                        if ($allele{$num1[$z]}{$markers_jrk[$k]} eq $allele{$num2[$z]}{$markers_jrk[$k]}){
                        #相同   (A,A || AB,AB)
                                $type[$z][$k] = '';
                                $count[$z][$k] = '';
                        }elsif ($markers_jrk[$k] eq 'Amel'){  # 性染色体上的marker不参与嵌合率计算
                                $type[$z][$k] = '';
                                $count[$z][$k] = '';
                        }elsif ($markers_jrk[$k] eq 'Yindel'){  # 性染色体上的marker不参与嵌合率计算. Update in v20200720
                                $type[$z][$k] = '';
                                $count[$z][$k] = '';
                        }elsif (@allele2 == 1 && @allele1 == 2 && ($allele1[0] eq $allele2[0] || $allele1[1] eq $allele2[0])){
                        #供者纯合&&供患有一个相同   (AB,A || AB,B)
                                $type[$z][$k] = 2;
                        }elsif (@allele1 == 1 && @allele2 == 2 && ($allele1[0] eq $allele2[0] || $allele1[0] eq $allele2[1])){
                        #患者纯合&&供患有一个相同   (A,AB || B,AB)
                                $type[$z][$k] = 3;
                        }elsif ((@allele1 == 2 && @allele2 == 2 && $allele1[0] ne $allele2[0] && $allele1[0] ne $allele2[1]  && $allele1[1] ne $allele2[0] && $allele1[1] ne $allele2[1]) ||@allele1 == 1 && @allele2 == 2 && $allele1[0] ne $allele2[0] && $allele1[0] ne $allele2[1] || @allele1 == 2 && @allele2 == 1 && $allele1[0] ne $allele2[0] && $allele1[1] ne $allele2[0] || @allele1 == 1 && @allele2 == 1 && $allele1[0] ne $allele2[0]){
                        #完全不同   (AB,CD || A,CD || AB,C || A,C )
                                $type[$z][$k] = 1;
                        }elsif (  @allele1 == 2 && @allele2 == 2 && (($allele1[0] eq $allele2[0] && $allele1[1] ne $allele2[1]) || ($allele1[1] eq $allele2[1] && $allele1[0] ne $allele2[0]) || ($allele1[1] eq $allele2[0] && $allele1[0] ne $allele2[1]) || ($allele1[0] eq $allele2[1] && $allele1[1] ne $allele2[0]))){
                        #均杂合&&有一个相同  (5 6,5 7 || 5 6,4 6 || 5 6,6 7 || 5 6,4 5)
                                $type[$z][$k] = 4;
                        }else{
                                $type[$z][$k] = "error";
                                $count[$z][$k] = "error";
                        }

                        print "L1996:Type: ",$type[$z][$k],"\n";

                        my %areas;
                        for my $p (0..$#allele3){
                                $areas{$allele3[$p]} = $area3[$p];
                        }  print "L2001:$z\t$k\t$markers_jrk[$k]\t$type[$z][$k]\n";
                        if ($type[$z][$k] eq 1){

                                if (@allele1 == 2 && @allele2 == 2){
                                        if (@allele3 == 2 && $allele1[0] eq $allele3[0] && $allele1[1] eq $allele3[1]){
                                        # AB,CD,AB
                                                $count[$z][$k] = 0;
                                        }elsif (@allele3 == 2 && $allele2[0] eq $allele3[0] && $allele2[1] eq $allele3[1]){
                                        # AB,CD,CD
                                                $count[$z][$k] = 1;
                                        }elsif (@allele3 == 2 && exists $areas{$allele2[0]} && !exists $areas{$allele2[1]}){
                                        # AB,CD,AC || AB,CD,BC
                                                $count[$z][$k] = $areas{$allele2[0]} / ($area3[0] + $area3[1]);
                                        }elsif (@allele3 == 2 && exists $areas{$allele2[1]} && !exists $areas{$allele2[0]}){
                                        # AB,CD,AD || AB,CD,BD
                                                $count[$z][$k] = $areas{$allele2[1]} / ($area3[0] + $area3[1]);
                                        }elsif (@allele3 == 3 && exists $areas{$allele2[0]} && !exists $areas{$allele2[1]}){
                                        # AB,CD,ABC
                                                $count[$z][$k] = $areas{$allele2[0]} / ($area3[0] + $area3[1] + $area3[2]);
                                        }elsif (@allele3 == 3 && exists $areas{$allele2[1]} && !exists $areas{$allele2[0]}){
                                        # AB,CD,ABD
                                                $count[$z][$k] = $areas{$allele2[1]} / ($area3[0] + $area3[1] + $area3[2]);
                                        }elsif (@allele3 == 3 && exists $areas{$allele2[0]} && exists $areas{$allele2[1]}){
                                        # AB,CD,ACD || AB,CD,BCD
                                                $count[$z][$k] = ($areas{$allele2[0]} + $areas{$allele2[1]}) / ($area3[0] + $area3[1] + $area3[2]);
                                        }elsif (@allele3 == 4 ){
                                        # AB,CD,ABCD
                                                $count[$z][$k] = ($areas{$allele2[0]} + $areas{$allele2[1]}) / ($area3[0] + $area3[1] + $area3[2] + $area3[3]);
                                        }else{
                                                $count[$z][$k] = "error";
                                        }
                                }elsif(@allele1 == 1 && @allele2 == 1){
                                        if(@allele3 == 1 && $allele1[0] eq $allele3[0]) {
                                        # A,B,A
                                                $count[$z][$k] = 0;
                                        }elsif (@allele3 == 1 && $allele2[0] eq $allele3[0]) {
                                        # A,B,B
                                                $count[$z][$k] = 1;
                                        }elsif (@allele3 == 2){
                                        # A,B,AB
                                                $count[$z][$k] = $areas{$allele2[0]} / ($area3[0] + $area3[1]);
                                        }else{
                                                $count[$z][$k] = "error";
                                        }
                                }elsif(@allele1 == 1 && @allele2 == 2){
                                        if(@allele3 == 1 && $allele1[0] eq $allele3[0]) {
                                        # A,BC,A
                                                $count[$z][$k] = 0;
                                        }elsif (@allele3 == 2 && $allele2[0] eq $allele3[0] && $allele2[1] eq $allele3[1]){
                                        # A,BC,BC
                                                $count[$z][$k] = 1;
                                        }elsif (@allele3 == 2 && exists $areas{$allele2[0]} && !exists $areas{$allele2[1]}){
                                        # A,BC,AB
                                                $count[$z][$k] = $areas{$allele2[0]} / ($area3[0] + $area3[1]);
                                        }elsif (@allele3 == 2 && exists $areas{$allele2[1]} && !exists $areas{$allele2[0]}){
                                        # A,BC,AC
                                                $count[$z][$k] = $areas{$allele2[1]} / ($area3[0] + $area3[1]);
                                        }elsif (@allele3 == 3 ){
                                        # A,BC,ABC
                                                $count[$z][$k] = ($areas{$allele2[0]} + $areas{$allele2[1]}) / ($area3[0] + $area3[1] + $area3[2]);
                                        }else{
                                                $count[$z][$k] = "error";
                                        }
                                }elsif(@allele1 == 2 && @allele2 == 1){
                                        if(@allele3 == 1 && $allele2[0] eq $allele3[0]) {
                                        # AB,C,C
                                                $count[$z][$k] = 1;
                                        }elsif (@allele3 == 2){
                                                if ($allele1[0] eq $allele3[0] && $allele1[1] eq $allele3[1]){
                                                # AB,C,AB
                                                        $count[$z][$k] = 0;
                                                }else{
                                                        $count[$z][$k] = $areas{$allele2[0]} / ($area3[0] + $area3[1]);
                                                        # AB,C,AC || AB,C,BC
                                                }
                                        }elsif (@allele3 == 3 ){
                                        # AB,C,ABC
                                                $count[$z][$k] = $areas{$allele2[0]} / ($area3[0] + $area3[1] + $area3[2]);
                                        }else{
                                                $count[$z][$k] = "error";
                                        }
                                }
                        }elsif ($type[$z][$k] eq 2){
                                $count[$z][$k] = "NA";

                        }elsif ($type[$z][$k] eq 3){
                                $count[$z][$k] = "NA";

                        }elsif ($type[$z][$k] eq 4){ print "L2089:$allele2[0]\t$allele3[0]\t $allele2[1]\t$allele3[1]\n";
                                if(@allele3 == 2 && $allele2[0] eq $allele3[0] && $allele2[1] eq $allele3[1]){
                                # AB,AC,AC
                                        $count[$z][$k] = 1;
                                }elsif (@allele3 == 2 && $allele1[0] eq $allele3[0] && $allele1[1] eq $allele3[1]){
                                # AB,AC,AB
                                        $count[$z][$k] = 0;
                                }elsif (@allele3 == 3){
                                # AB,AC,ABC
                                        if($allele1[0] eq $allele2[0] && $allele1[1] ne $allele2[1]){
                                        # 5 6,5 7,5 6 7
                                                $count[$z][$k] = $areas{$allele2[1]}/($areas{$allele2[1]}+$areas{$allele1[1]});
                                        }elsif ($allele1[1] eq $allele2[1] && $allele1[0] ne $allele2[0]){
                                        # 5 6,4 6,4 5 6
                                                $count[$z][$k] = $areas{$allele2[0]}/($areas{$allele2[0]}+$areas{$allele1[0]});
                                        }elsif ($allele1[1] eq $allele2[0] && $allele1[0] ne $allele2[1]){
                                        # 5 6,6 7,5 6 7
                                                $count[$z][$k] = $areas{$allele2[1]}/($areas{$allele2[1]}+$areas{$allele1[0]});
                                        }elsif ($allele1[0] eq $allele2[1] && $allele1[1] ne $allele2[0]){
                                        # 5 6,4 5,4 5 6
                                                $count[$z][$k] = $areas{$allele2[0]}/($areas{$allele2[0]}+$areas{$allele1[1]});
                                        }else{
                                                $count[$z][$k] = "error";
                                        }
                                }else{
                                        $count[$z][$k] = "error";
                                }
                        }
                }
                # <STDIN>;

                # 判断 报告单编号 的结论是否为空
                if ($conclusion[$z]){  # 不为空
                        # 跳过 结论为 '跳过' 的报告单
                        if ($conclusion[$z] eq '跳过'){
                                next;
                        }
                }

                ################################################  2020.05.22.16.18 ################################################
                my @temp_marker = ();  # 记录每个marker 是 '混合嵌合' 还是 ' '（完全嵌合）
                my @tempcount = ();  # 记录每份报告 有嵌合率结果位点 的嵌合率
                foreach my $k (0..$#markers_jrk){ print "L2131:$z\t$k\t$count[$z][$k]\n";
                        if ($count[$z][$k] =~ /\d/){
                                # $count_sum[$z] += $count[$z][$k];
                                # $count_n[$z] += 1;
                                push @tempcount, $count[$z][$k];
                                if ($count[$z][$k]<1 && $count[$z][$k]>0){  # 位点的嵌合率在 0-1之间，定义为 '混合嵌合'
                                        $temp_marker[$k] = '混合嵌合';
                                }else{
                                        $temp_marker[$k] = ' ';
                                }
                        }else{
                                $temp_marker[$k] = ' ';
                        }
                }

                foreach my $k (0..$#markers_jrk){
                        $marker_type[$z][$k] = $temp_marker[$k];  # 记录 第几个报告单编号 的第几个marker 的状态：'混合嵌合' / ' ' (完全嵌合)
                }
                $count_n[$z] = scalar(@tempcount);  # 记录每份报告 有嵌合率结果的位点个数

                if ($count_n[$z] > 0){  # 该报告单编号 对应的有效位点（有嵌合率结果的位点）> 0
                        ($count_avg[$z], $SD[$z]) = &Avg_SD(@tempcount);  # 计算所有嵌合位点，嵌合率的均值 和 SD，分别存入 $count_avg[$z], $SD[$z]
                        $count_avg[$z] = sprintf("%.4f", $count_avg[$z]);
                }else{  # 该报告单编号 没有 有效位点，报错，提示无法出具此份报告，跳过该报告单
                        $success = 0;
                        $error = '报告单号'.$TCAID.'没有有效位点，请检查！将跳过出具此份报告。';
                        Win32::MsgBox $error, 0, "注意";
                        $RptBox -> Append("失败【无有效位点】\r\n");
                        $conclusion[$z] = '跳过';
                        next;
                }

                # if ($count_n[$z] != 0){
                        # $count_avg[$z] = $count_sum[$z] / $count_n[$z];
                        # $count_avg[$z] = sprintf("%.4f", $count_avg[$z]);
                # }

                $RptBox -> Append("成功！\r\n");  # 更新 "生成报告" 部分的文本框中显示的提示信息 "成功！"
                next if $exp_num{$TCAID} != 3;  # 如果该 报告单编号 对应的实验编码个数 不为3，则跳过将本次实验的结果保存在内存中，跳到下一个报告单编号
                ##追加本次实验结果到内存中
                next if $count_avg[$z] !~ /\d/;  # 如果该 报告单编号 对应的平均嵌合率 不是数字（什么情况？），则跳过将本次实验的结果保存在内存中，跳到下一个报告单编号
                my $tempid = $identity{$TCAID};  # 获取 报告单编号 对应的 患者编码 # %identity 存储 报告单编号 <=> 患者编码(HUN001胡琳)
                push @{$Chimerism{$tempid}}, sprintf ("%.2f%s", $count_avg[$z]*100,"%");  # 将该 患者编码 对应的嵌合结果（有效位点嵌合率的平均值）存入$Chimerism{$tempid}
                push @{$SampleID{$tempid}}, $num3[$z];  # 记录该 患者编码 对应的 "术后患者" 实验编码，存入 $SampleID{$tempid}
                push @{$ReportDate{$tempid}}, sprintf ("%d-%02d-%02d", $year, $mon, $mday);  # 记录该 患者编码 对应的 报告日期，存入 $ReportDate{$tempid}
                if ($cells{$num3[$z]} ne "-"){  # "术后患者" 的 "分选类型" 不为 "-"
                        $sampleType{$num3[$z]} = $cells{$num3[$z]};  # 将 "术后患者" 实验编码 对应的 "样本类型" 设置为 供患信息中 实验编码 对应的 分选类型
                }else{  # "术后患者" 的 "分选类型" 为 "-"
                        $sampleType{$num3[$z]} = $sample{$num3[$z]};  # 将 "术后患者" 实验编码 对应的 "样本类型" 设置为 供患信息中 实验编码 对应的 样本类型
                }
                $receiveDate{$num3[$z]} = DateUnify($date1{$num3[$z]});  # 将 "术后患者" 实验编码 对应的 "收样日期" 设置为 供患信息中 实验编码 对应的 生产时间  （用生产日期 设置 收样日期 ？）
                $sampleDate{$num3[$z]} = DateUnify($date3{$num3[$z]});  # 将 "术后患者" 实验编码 对应的 "采样日期" 设置为 供患信息中 实验编码 对应的 采样日期
                ##坐等追加到总表中
                print "L2167:\$tempid:". $tempid,"\n";
                print "L2168:"."Chimerism:";print $_,"|" foreach (@{$Chimerism{$tempid}});print "\n";
                print "L2169:"."SampleID:";print $_,"|" foreach (@{$SampleID{$tempid}});print "\n";
                print "L2170:"."ReportDate:";print $_,"|" foreach (@{$ReportDate{$tempid}});print "\n";
                print "L2171:"."receiveDate:";print $_,"|" foreach (@{$ReportDate{$tempid}});print "\n";
                print "L2172:"."ReportDate:";print $_,"|" foreach (@{$ReportDate{$tempid}});print "\n";
        }
        my $chimerismSummary = sprintf "嵌合率汇总-%4d%02d%02d.txt",$year, $mon, $mday;  # 重新 写一个 嵌合率汇总 的结果文件
        my $ana_date = sprintf "%4d-%02d-%02d",$year, $mon, $mday; print "ana_date: $ana_date\n";
        open SUM,"> $chimerismSummary";
        print SUM "姓名\t医院\t样本类型\t样本编号\t报告编号\t嵌合率\t有效位点\tSD\tCV\n";

        ################################################  2020.05.22.17.04 ################################################
        $RptBox -> Append("输出准备完成！开始输出报告\r\n========================\r\n");  # 更新 "生成报告" 部分的文本框中显示的提示信息
        # 遍历输出每一份报告
        foreach my $z(0..$#TCA_id){
                my $TCAID = $TCA_id[$z];  print "L2180:$z|$TCAID|$conclusion[$z]\n" ;# 获取 报告单编号
                $RptBox -> Append($TCAID.'...');  # 更新 "生成报告" 部分的文本框中显示的提示信息
                if ($conclusion[$z] eq '跳过'){  # 跳过 结论为 "跳过" 的报告单
                        $RptBox -> Append("跳过\r\n");
                        next;
                }
                # %sheet_name:
                # 一批检测中 同一个 "患者姓名" 按照 CSTB 的设计，会对应多份 报告单(即, 对应多个 报告单编号)

                # 为每个检测样本新建一个结果输出目录
                my $tempid = $identity{$TCAID};  # 获取 报告单编号 对应的 患者编码 # %identity 存储 报告单编号 <=> 患者编码(HUN001胡琳)
                print "L2215:$Output_Dir\\$name{$num3[$z]}-STR检测报告-$hospital{$num3[$z]}-$ana_date\n";
                if(-d "$Output_Dir\\$name{$num3[$z]}-STR检测报告-$hospital{$num3[$z]}-$ana_date"){
                    print "L2217:Dir existed:$Output_Dir\\$name{$num3[$z]}-STR检测报告-$hosptl_num{$num3[$z]}-$ana_date\n" ;
                } else {
                    print "L2220:CreateDir:$Output_Dir\\$name{$num3[$z]}-STR检测报告-$hospital{$num3[$z]}-$ana_date\n" ;
                    Win32::CreateDirectory("$Output_Dir\\$name{$num3[$z]}-STR检测报告-$hospital{$num3[$z]}-$ana_date");  # 新建文件夹, 带有中文名
                }

                # 定义 报告单 输出文件
                if (exists $sheet_name{$sheet[$z]}){  # @sheet: 存放 第几份 报告对应的 "患者姓名"
                        $sheet_name{$sheet[$z]} += 1;
                        $Output_rpt_str = sprintf "%s\\%s-STR检测报告-%s-%s\\%s-%s%s%d.xlsx", $Output_Dir, $name{$num3[$z]}, $hospital{$num3[$z]}, $ana_date, $TCAID, $sheet[$z], 'AK', $sheet_name{$sheet[$z]};
                }else{
                        $Output_rpt_str = sprintf "%s\\%s-STR检测报告-%s-%s\\%s-%s%s.xlsx", $Output_Dir, $name{$num3[$z]}, $hospital{$num3[$z]}, $ana_date, $TCAID, $sheet[$z], 'AK';
                        $sheet_name{$sheet[$z]} = 1;
                }

                ################################################  2020.05.22.17.28 ################################################
                my $workbook;
                # 打开 报告输出文件，准备写报告
                unless ($workbook = Excel::Writer::XLSX->new($Output_rpt_str)){
                        $error = $Output_rpt_str."无法保存！";
                        Win32::MsgBox $error, 0, "错误！";
                        $success = 0;
                        $RptBox -> Append($Output_rpt_str."打开失败！跳过\r\n");
                        next;
                }

                ###################  定义excel文件各部分的格式  #######################################################
                my $format1  = $workbook->add_format(size => 18, bold => 1, align => 'center',                      font => decode('GB2312','楷体')); # HLA高分辨基因分型检测报告
                my $format2  = $workbook->add_format(size => 11,                                                                                     'top' => 1, 'bottom' => 2);  # 双线
                my $format3  = $workbook->add_format(size => 11,            align => 'right',  valign => 'vcenter', font => decode('GB2312','宋体')); # 报告单编号
                my $format4  = $workbook->add_format(size => 12, bold => 1, align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 送检单位，检测项目 write
                my $format5  = $workbook->add_format(size => 12, bold => 1, align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 送检单位，检测项目 merge
                my $format6  = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 信息/结果，宋体，write
                my $format7  = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 信息/结果，宋体，merge
                my $format8  = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => 'Times New Roman',       'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 信息/结果，Times New Roman，write
                my $format9  = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => 'Times New Roman',       'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 信息/结果，Times New Roman，merge
                my $format10 = $workbook->add_format(size => 10,            align => 'center', valign => 'vcenter', font => 'Times New Roman',       'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 信息/结果，Times New Roman，write，字号10
                #my $format11 = $workbook->add_format(size => 10,                               valign => 'vcenter', font => decode('GB2312','宋体'), text_wrap => 1); # 备注
                #my $format12 = $workbook->add_format(size => 11, bold => 1,                    valign => 'vcenter', font => decode('GB2312','宋体')); #
                #my $format13 = $workbook->add_format(size => 11, bold => 1,                    valign => 'vcenter', font => decode('GB2312','宋体'),             'bottom' => 1);  # 检测者
                #my $format14 = $workbook->add_format(size => 11, bold => 1,                    valign => 'vcenter', font => 'Times New Roman',                   'bottom' => 1);  # 报告日期
                my $format15 = $workbook->add_format(size => 12,            align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 'STR位点' '位点状态' '备注'
                my $format16 = $workbook->add_format(size => 9,             align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 结果栏样本编号
                my $format17 = $workbook->add_format(size => 12, bold => 1,                    valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 0); # '检测结论'
                my $format18 = $workbook->add_format(size => 11,                               valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 0, 'right' => 1); # 检测结论
                my $format19 = $workbook->add_format(size => 8, valign => 'vcenter', font => decode('GB2312','华文中宋'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1, 'text_wrap' => 1,); #备注

                #########################  chart的格式 ############################
                my $Gfmt1 = $workbook->add_format(size => 10, align => 'right', font => decode('GB2312','宋体'));  # chart患者姓名
                my $Gfmt2 = $workbook->add_format(size => 14, bold => 1, align => 'center', font => decode('GB2312','宋体')); # chart姓名
                my $Gfmt3 = $workbook->add_format(size => 12, bold => 1, align => 'center', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1);# chart表头
                my $Gfmt4 = $workbook->add_format(size => 10, align => 'center', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # chart表
                my $Gfmt5 = $workbook->add_format(size => 14, bold => 1, font => decode('GB2312','宋体'));  # TCA定期检测流程
                my $Gfmt6 = $workbook->add_format(size => 11, font => decode('GB2312','宋体'));  # 提示
                my $Gfmt7 = $workbook->add_format(size => 11, align => 'center', font => decode('GB2312','宋体')); # 温馨提示
                my $Gfmt8 = $workbook->add_format(size => 10, align => 'center', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # chart表
                my $Gfmt9 = $workbook->add_format(size =>  9, align => 'center', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # chart表
                $Gfmt8->set_num_format('0.00');
                #####

                my ($countsheet, $graphic, $worksheet, $graphic_temp);
                $worksheet  = $workbook ->add_worksheet(decode('GB2312',"报告"));
                # 同一个 "患者姓名" 仅在第一份报告单中 输出 "嵌合曲线" 和 "tmp" 表单
                if ($exp_num{$TCAID} == 3 && $sheet_name{$sheet[$z]} == 1){  # 报告单编号 对应 3个 实验编码，同时 $sheet[$z] 对应的 "患者姓名" 的第一份报告
                        $graphic = $workbook ->add_worksheet(decode('GB2312',"嵌合曲线"));
                }
                $countsheet = $workbook->add_worksheet(decode('GB2312',"计算"));
                if ($exp_num{$TCAID} == 3 && $sheet_name{$sheet[$z]} == 1){  # 报告单编号 对应 3个 实验编码，同时 $sheet[$z] 对应的 "患者姓名" 的第一份报告
                        $graphic_temp = $workbook->add_worksheet('temp');
                }

                ######################################## 输出 "计算" 表单 ########################################
                $countsheet->hide_gridlines();  # "计算" 表单中，隐藏网格线
                $countsheet->keep_leading_zeros();  # "计算"表单中，保留数字开头的0
                # 定义 "计算" 表单里的单元格格式
                my $format101 = $workbook->add_format(size  => 11, font  => decode('GB2312','宋体'));
                my $format102 = $workbook->add_format(size  => 11, align => 'center', font  => decode('GB2312','宋体'));
                my $format103 = $workbook->add_format(size  => 11, color => 'red', font   => decode('GB2312','宋体'));

                # 写 "计算" 表单
                # 写标题
                $countsheet->write('A01',decode('GB2312','位点'), $format101);
                $countsheet->merge_range('B1:C1', decode('GB2312','患者'), $format102);
                $countsheet->merge_range('D1:E1', decode('GB2312','供者'), $format102);
                $countsheet->merge_range('F1:I1', decode('GB2312','术后'), $format102);
                $countsheet->write('J01',decode('GB2312','类型'), $format101);
                $countsheet->merge_range('K1:N1', decode('GB2312','面积'), $format102);
                $countsheet->merge_range('B2:C2', decode('GB2312',$num1[$z]),  $format102);
                $countsheet->merge_range('D2:E2', decode('GB2312',$num2[$z]),  $format102);
                $countsheet->merge_range('F2:I2', decode('GB2312',$num3[$z]),  $format102);
                $countsheet->merge_range('O1:S1', decode('GB2312','嵌合率'),  $format102);
                $countsheet->write('O2', 'TYPE1',  $format101);
                $countsheet->write('P2', 'TYPE2',  $format101);
                $countsheet->write('Q2', 'TYPE3',  $format101);
                $countsheet->write('R2', 'TYPE4',  $format101);
                $countsheet->write('S2', 'ERROR',  $format101);
                $countsheet->write('T1',decode('GB2312','总嵌合率'), $format101);

                ################################################  2020.05.24.14.32 ################################################
                # 从第一列的三行开始，将 marker名 写入第一列
                for my $j (0..$#markers_jrk){
                        $countsheet->write($j+2,0,$markers_jrk[$j], $format101);
                }
                # 遍历写入每个 marker "术前患者" "术前供者" "术后患者" 样本的 型别(%allele),面积(%area),状态(%type),嵌合率(%count)
                foreach my $k (0..$#markers_jrk){
                        my @allele1 = split/,/, $allele{$num1[$z]}{$markers_jrk[$k]};
                        my @allele2 = split/,/, $allele{$num2[$z]}{$markers_jrk[$k]};
                        my @allele3 = split/,/, $allele{$num3[$z]}{$markers_jrk[$k]};
                        my @area3   = split/,/, $area{$num3[$z]}{$markers_jrk[$k]};
                        for my $l (0..$#allele1){
                                $countsheet->write($k+2,$l+1,$allele1[$l], $format101);
                        }
                        for my $m (0..$#allele2){
                                $countsheet->write($k+2,$m+3,$allele2[$m], $format101);
                        }
                        for my $n (0..$#allele3){
                                $countsheet->write($k+2,$n+5,$allele3[$n], $format101);
                        }
                        for my $o (0..$#area3){
                                $countsheet->write($k+2,$o+10,$area3[$o], $format101);
                        }

                        if ($type[$z][$k] ne 'error'){
                                $countsheet->write($k+2,9,$type[$z][$k],$format101);
                        }else{
                                $countsheet->write($k+2,9,$type[$z][$k],$format103);
                        }

                        if ($count[$z][$k] =~ /\d/ && $count[$z][$k] <= 1 && $count[$z][$k] >= 0){
                                if ($type[$z][$k] =~ /\d/){
                                        $countsheet->write($k+2,13+$type[$z][$k],$count[$z][$k],$format101);
                                }else{
                                        $countsheet->write($k+2,18,$count[$z][$k],$format103);
                                }
                        }else{
                                if ($type[$z][$k] =~ /\d/){
                                        $countsheet->write($k+2,13+$type[$z][$k],$count[$z][$k],$format103);
                                }else{
                                        $countsheet->write($k+2,18,$count[$z][$k],$format103);
                                }
                        }
                }

                ########################################### 输出 "报告" 表单 ##########################################################
                # 定义 "报告" 表单的各种格式
                $worksheet->hide_gridlines();
                $worksheet->keep_leading_zeros();
                # 设置 列的宽度、行的高度
                $worksheet->set_column(0,0,0.5);
                $worksheet->set_column(1,1,14.5);
                $worksheet->set_column(2,2,10);
                $worksheet->set_column(3,4,8.5);
                $worksheet->set_column(5,5,11);
                $worksheet->set_column(6,7,10.5);
                $worksheet->set_column(8,8,14);
                my @rows = (73,8,3,18.4,22.8,22.8,10,16.2,16.2,16.2,16.2,16.2,10,18.6,18.6,18.6,18.6,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,39,57.6,12.6,25,25);
                for my $i (0 .. $#rows){$worksheet->set_row($i, $rows[$i]);}

                # 设置页面的左右及上边距
                $worksheet->set_margin_left(0.394);
                $worksheet->set_margin_right(0.394);
                $worksheet->set_margin_top(0.2);

                # 设置表单的 页脚
                #my $footer = '&R'.decode('GB2312',$name{$num3[$z]}.'-'.$hospital{$num3[$z]}.'（'.$doctor{$num3[$z]}.'），第').'&P'.decode('GB2312','页/共').'&N'.decode('GB2312','页');
                # my $footer = '&L'.decode('GB2312','检验实验室：荻硕贝肯检验实验室')."\n".
                             #'&R'.decode('GB2312','CSTB-B-R-0021-1.0')."\n".
                             '&L'.decode('GB2312','咨询电话：020-62313880').
                             '&R'.decode('GB2312',$name{$num3[$z]}.'，第').'&P'.decode('GB2312','页/共').'&N'.decode('GB2312','页');

                # $worksheet->set_footer($footer);

                # 在第一行第二列插入 公司logo (pic/logo.png)
                $worksheet->insert_image('B1', "pic/logo.png", 10, 10, 0.15, 0.15);

                # 写入 各项信息
                $worksheet->merge_range('B1:I1', decode('GB2312','嵌合状态分析报告'), $format1);
                $worksheet->merge_range('B3:I3', decode('GB2312',''), $format2);
                $worksheet->merge_range('G4:I4', decode('GB2312','报告单号：'.$rptnum{$num3[$z]}.'   '),$format3);
                $worksheet->write('B05',decode('GB2312','送检单位'),$format4);
                $worksheet->write('B06',decode('GB2312','检测项目'),$format4);
                $worksheet->merge_range('C5:G5', decode('GB2312',$hospital{$num3[$z]}),$format5);
                $worksheet->write('H05',decode('GB2312','送检医生'),$format4);
                if ($doctor{$num3[$z]} =~ /^([\x80-\xFF][\x80-\xFF])$/){   #单字 中文
                     $doctor{$num3[$z]} = $1.'医生';
                     }
                $worksheet->write('I05',decode('GB2312',$doctor{$num3[$z]}),$format4);
                my $testt;
                my $cell_type = "" ;
                if ($cells{$num3[$z]} =~ /(\S+)细胞分选/){  # 供患信息中 实验编码 对应的 分选类型 (B细胞分选 / T细胞分选 / NK细胞分选 ...)
                    $cell_type = $1 ;  # B / T / NK / 粒
                    $testt = $cell_type.'细胞嵌合状态分析';  # 根据 "分选类型" 动态输出 "检测项目"
                }
                else{
                    $testt = '全血嵌合状态分析';
                }
                $worksheet->merge_range('C6:I6', decode('GB2312',$testt),$format5);  # 写入 "检测项目" 内容
                $worksheet->merge_range('B8:I8', decode('GB2312','样本信息'),$format7);  # 写入 "样本信息" 表头
                $worksheet->write('B09',decode('GB2312','样本编号'),$format6);  # 写入 "样本编号" 表头
                $worksheet->write('C09',decode('GB2312','姓名'),$format6);  # 写入 "姓名" 表头
                $worksheet->write('D09',decode('GB2312','性别'),$format6);  # 写入 "型别" 表头
                $worksheet->write('E09',decode('GB2312','年龄'),$format6);  # 写入 "年龄" 表头
                $worksheet->write('F09',decode('GB2312','样本类型'),$format6);  # 写入 "样本类型" 表头
                $worksheet->write('G09',decode('GB2312','采样日期'),$format6);  # 写入 "采样日期" 表头
                $worksheet->write('H09',decode('GB2312','收样日期'),$format6);  # 写入 "收样日期" 表头
                $worksheet->write('I09',decode('GB2312','关系'),$format6);  # 写入 "关系" 表头
                $worksheet->write('B10',decode('GB2312',$number{$num3[$z]}),$format10);  # 写入 "术后患者" 实验编码 (报告中标题是 "样本编号"，最好统一起来)
                $worksheet->write('B11',decode('GB2312',$number{$num2[$z]}),$format10);  # 写入 "术前供者" 实验编码
                $worksheet->write('C10',decode('GB2312',$name{$num3[$z]}),$format6);  # 写入 "术后患者" 姓名
                $worksheet->write('C11',decode('GB2312',$name{$num2[$z]}),$format6);  # 写入 "术前供者" 姓名
                $worksheet->write('D10',decode('GB2312',$gender{$num3[$z]}),$format6);  # 写入 "术后患者" 性别
                $worksheet->write('D11',decode('GB2312',$gender{$num2[$z]}),$format6);  # 写入 "术前供者" 性别
                $worksheet->write('E10',decode('GB2312',$age{$num3[$z]}),$format6);  # 写入 "术后患者" 年龄
                $worksheet->write('E11',decode('GB2312',$age{$num2[$z]}),$format6);  # 写入 "术前供者" 年龄
                $worksheet->write('F10',decode('GB2312',$sample{$num3[$z]}),$format6);  # "术后患者" 在供患信息中 实验编码 对应的 样本类型
                $worksheet->write('F11',decode('GB2312',$sample{$num2[$z]}),$format6);  # "术前供者" 在供患信息中 实验编码 对应的 样本类型
                $worksheet->write('G10',decode('GB2312',DateUnify($date3{$num3[$z]})),$format10);  # "术后患者" 的采样日期 设置为 供患信息中 实验编码 对应的 采样日期
                $worksheet->write('G11',decode('GB2312',DateUnify($date3{$num2[$z]})),$format10);  # "术前供者" 的采样日期 设置为 供患信息中 实验编码 对应的 采样日期
                $worksheet->write('H10',decode('GB2312',DateUnify($date4{$num3[$z]})),$format10);  # "术后患者" 的收样日期 设置为 供患信息中 实验编码 对应的 收样日期
                $worksheet->write('H11',decode('GB2312',DateUnify($date4{$num2[$z]})),$format10);  # "术前供者" 的收样日期 设置为 供患信息中 实验编码 对应的 收样日期
                $worksheet->write('B12',decode('GB2312','住院/门诊号'),$format6);  # 写入 "住院/门诊号" 表头
                $worksheet->write('E12',decode('GB2312','床号'),$format6);  # 写入 "床号" 表头
                $worksheet->write('G12',decode('GB2312','临床诊断'),$format6);  # 写入 "临床诊断" 表头
                $worksheet->merge_range('C12:D12', decode('GB2312',$hosptl_num{$num3[$z]}), $format7);  # 患者的 "住院/门诊号" 设置为 供患信息中 "术后患者" 实验编码 对应的 住院号
                $worksheet->write('F12',decode('GB2312',$bed_num{$num3[$z]}),$format6);  # 患者的 "床号" 设置为 供患信息中 "术后患者" 实验编码 对应的 床号

                # 写入 患者的 临床诊断 信息
                if ($diagnosis{$num3[$z]} ne "-"){  # "术后患者" 的临床诊断不为 "-"
                     $worksheet->merge_range('H12:I12', decode('GB2312',$diagnosis{$num3[$z]}), $format7);  # 根据 供患信息中 "诊断" 设置
                }else{
                     $worksheet->merge_range('H12:I12', decode('GB2312',$diagnosis{$num1[$z]}), $format7);  # 根据术前患者的 实验编码 对应 供患信息中 "诊断" 设置
                }

                my $tmp = $sheet[$z];  # @sheet: 存放 第几份 报告对应的 "患者姓名"
                if ($relation{$num2[$z]} =~ /$tmp/){  # %relation 供患信息中 实验编码 对应的 亲缘关系 / $num2[$z]: 第几份报告 对应的 "术前供者" 的实验编码
                        $relation{$num2[$z]} =~ s/$tmp//;  # 如: 李某父亲， 去掉其中的 "李某"
                }
                $worksheet->write('I10',decode('GB2312',$relation{$num3[$z]}),$format6);  # 写入 "术后患者" 与 患者的关系
                $worksheet->write('I11',decode('GB2312',$relation{$num2[$z]}),$format6);  # 写入 "术前供者" 与 患者的关系
                $worksheet->merge_range('B14:I14', decode('GB2312','检测结果'), $format5);  # 写入 "检测结果" 表头
                $worksheet->merge_range('B15:B17', decode('GB2312','STR位点'), $format15);  # 写入 "STR位点" 表头
                $worksheet->merge_range('C15:H15', decode('GB2312','等位基因'), $format5);  # 写入 "等位基因" 表头
                if ($sample{$num1[$z]} =~ /口腔/){  # "术前患者" 在供患信息中 实验编码 对应的 样本类型
                        $worksheet->merge_range('C16:D16', decode('GB2312','患者移植前(口腔)'), $format7);  # 写入 "患者移植前(口腔)" 表头
                }else{
                        $worksheet->merge_range('C16:D16', decode('GB2312','患者移植前'), $format7);  # 写入 "患者移植前" 表头
                }
                $worksheet->merge_range('E16:F16', decode('GB2312','供    者'), $format7);  # 写如 "供者" 表头
                $worksheet->merge_range('G16:H16', decode('GB2312','患者移植后'), $format7);  # 写入 "患者移植后" 表头
                $worksheet->merge_range('C17:D17', decode('GB2312','样本编号：'.$num1[$z]), $format16);  # 写入 "术前患者" 的 实验编码
                $worksheet->merge_range('E17:F17', decode('GB2312','样本编号：'.$num2[$z]), $format16);  # 写入 "术前供者" 的 实验编码
                $worksheet->merge_range('G17:H17', decode('GB2312','样本编号：'.$num3[$z]), $format16);  # 写入 "术后患者" 的 实验编码
                $worksheet->merge_range('I15:I17', decode('GB2312','位点状态'), $format15);  # 写入 "位点状态" 表头
                # 写入 每个marker的型别
                for my $q (0..$#markers_jrk){
                        $worksheet->write($q+17,1,$markers_jrk[$q], $format6);  # 写入 marker 名
                        $worksheet->merge_range($q+17,2,$q+17,3,decode('GB2312',$allele{$num1[$z]}{$markers_jrk[$q]}), $format7); # 写入 "术前患者" 每个marker 对应的型别
                        $worksheet->merge_range($q+17,4,$q+17,5,decode('GB2312',$allele{$num2[$z]}{$markers_jrk[$q]}), $format7);  # 写入 "术前供者" 每个marker 对应的型别
                        $worksheet->merge_range($q+17,6,$q+17,7,decode('GB2312',$allele{$num3[$z]}{$markers_jrk[$q]}), $format7);  # 写入 "术后患者" 每个marker 对应的型别
                        $worksheet->write($q+17,8,decode('GB2312',$marker_type[$z][$q]), $format6);  # 写入 第几个报告单编号 的第几个marker 的状态：'混合嵌合' / ' ' (完全嵌合)
                }

                # 写入 "检测结论" 表头
                $worksheet->write('B43',decode('GB2312','检测结论：'),$format17);
                # 写入 "检测结论"
                # 判断 样本的嵌合率 是否为 数字
                if ($count_avg[$z] =~ /\d/){  # 如果为数字，则转换为百分比，保留2位有效数字
                    print "L2448:$z|$count_avg[$z]|$conclusion[$z]\n";
                        $count_avg[$z] = sprintf("%.2f", $count_avg[$z]*100);
                        unless ($conclusion[$z]){  # 报告对应的结论 为 空
                                if($count_avg[$z] >= 95){  # 嵌合率 >= 95%  ==> 完全嵌合
                                        $conclusion[$z] = '患者移植后'. $sample{$num3[$z]} . '中供者'. $cell_type . '细胞占'.$count_avg[$z].'%，表现为完全嵌合状态。';
                                        $worksheet->merge_range('C43:I43',decode('GB2312',$conclusion[$z]), $format18);
                                }elsif($count_avg[$z] < 5){  # 嵌合率 < 5% ==> 微嵌合
                                        $conclusion[$z] = '患者移植后'. $sample{$num3[$z]} . '中供者'. $cell_type. '细胞占'.$count_avg[$z].'%，表现为微嵌合状态。';
                                        $worksheet->merge_range('C43:I43',decode('GB2312',$conclusion[$z]), $format18);
                                }else{  # 5% <= 嵌合率 < 95%  ==> 混合嵌合
                                        $conclusion[$z] = '患者移植后'. $sample{$num3[$z]} . '中供者'. $cell_type. '细胞占'.$count_avg[$z].'%，表现为混合嵌合状态。';
                                        $worksheet->merge_range('C43:I43',decode('GB2312',$conclusion[$z]), $format18);
                                }
                        }else{
                                $worksheet->merge_range('C43:I43',decode('GB2312',$conclusion[$z]), $format18);
                        }
                }else{
                        $worksheet->merge_range('C43:I43',decode('GB2312','无'), $format18);
                }

                ######################################### 存储 汇总报告单中需要的信息 ##########################################################
                 #################### 定义一组变量用于存储 汇总报告单中需要的信息 ########################################################################################
                #        my %this_patient_ID_and_report_id = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者本轮检测的汇总报告单的编号
                #        my %this_patient_ID_and_patient_name = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者姓名
                #        my %this_patient_ID_and_patient_gender = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者性别
                #        my %this_patient_ID_and_patient_age = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者年龄
                #        my %this_patient_ID_and_patient_diagnosis = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者 诊断信息
                #        my %this_patient_ID_and_patient_sampleType = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者样本类型 (外周血 / 骨髓 /骨髓血?)
                #        my %this_patient_ID_and_patient_sampleDetailType = () ;  # 定义一个hash表，用户存储 患者编码 (如:HUN001胡琳) 对应的 患者样本详细类型  (骨髓血-B细胞分选...)
                #        my %this_patient_ID_and_patient_sampleDate = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者采样日期
                #        my %this_patient_ID_and_patient_rcvDate = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 患者收样(接样)日期
                #        my %this_patient_ID_and_donor_name = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 供者姓名
                #        my %this_patient_ID_and_donor_gender = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 供者性别
                #        my %this_patient_ID_and_donor_age = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 供者年龄
                #        my %this_patient_ID_and_donor_relationship = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 供者与他/她的关系
                #        my %this_patient_ID_and_hospital = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 医院全称
                #        my %this_patient_ID_and_doctor = () ;  # 定义一个hash表，用于存储 患者编码 (如:HUN001胡琳) 对应的 送样医生
                #        my %this_patient_ID_sampleDetailType_and_conclusion = () ;  # 定义一个hash表，用于存储 患者编码 某个样本详细类型 对应的 结论
                #        my %this_patient_ID_and_shuqian_donor_expid = () ;  # 定义一个hash表，用于存储 患者编码 对用的 术前供者 实验编码
                #        my %this_patient_ID_and_shuqian_patient_expid = () ;  # 定义一个hash表，用于存储 患者编码 对用的 术前患者 实验编码
                #        my %this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_expid = () ;  # 定义一个hash表，用于存储 患者编码 对用的 术后患者(外周血/骨髓血) 实验编码
                #        my %this_patient_ID_and_shuhou_patient_T_cell_expid = () ;  # 定义一个hash表，用于存储 患者编码 对用的 术后患者(T细胞分选) 实验编码
                #        my %this_patient_ID_and_shuhou_patient_B_cell_expid = () ;  # 定义一个hash表，用于存储 患者编码 对用的 术后患者(B细胞分选) 实验编码
                #        my %this_patient_ID_and_shuhou_patient_NK_cell_expid = () ;  # 定义一个hash表，用于存储 患者编码 对用的 术后患者(NK细胞分选) 实验编码
                #        my %this_patient_ID_and_shuhou_patient_li_cell_expid = () ;  # 定义一个hash表，用于存储 患者编码 对用的 术后患者(粒细胞分选) 实验编码
                #        my %hash_this_patient_ID_and_shuqian_patient_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术前患者 的型别结果
                #        my %hash_this_patient_ID_and_shuqian_donor_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术前供者 的型别结果
                #        my %hash_this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术后(外周血/骨髓血) 的型别结果
                #        my %hash_this_patient_ID_and_shuhou_patient_T_cell_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术后(T细胞) 的型别结果
                #        my %hash_this_patient_ID_and_shuhou_patient_B_cell_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术后(B细胞) 的型别结果
                #        my %hash_this_patient_ID_and_shuhou_patient_NK_cell_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术后(NK细胞) 的型别结果
                #        my %hash_this_patient_ID_and_shuhou_patient_li_cell_genotypes = () ;  # 定义一个数组，用于存储 患者编码 对应的 术后(粒细胞) 的型别结果
                $tempid = $identity{$TCAID};  # 获取 报告单编号 对应的 患者编码 # %identity 存储 报告单编号 <=> 患者编码(HUN001胡琳)
                print "L2485:$z|$TCAID|$tempid\n" ;
                if(!exists $this_patient_ID_and_report_id{$tempid}){
                    my $report_id = (split(/-/,$TCAID))[0] ;
                    print "L2488|$tempid|report_id:$TCAID=>$report_id\n";
                    $this_patient_ID_and_report_id{$tempid} = $report_id ;
                }
                if(!exists $this_patient_ID_and_patient_name{$tempid}){
                    print "L2492|$tempid|patient_name$name{$num3[$z]}\n";
                    $this_patient_ID_and_patient_name{$tempid} = $name{$num3[$z]} ;
                }
                if(!exists $this_patient_ID_and_patient_gender{$tempid}){
                    print "L2496|$tempid|patient_gender:$gender{$num3[$z]}\n";
                    $this_patient_ID_and_patient_gender{$tempid} = $gender{$num3[$z]} ;
                }
                if(!exists $this_patient_ID_and_patient_age{$tempid}){
                    print "L2500|$tempid|patient_age:$age{$num3[$z]}\n";
                    $this_patient_ID_and_patient_age{$tempid} = $age{$num3[$z]} ;
                }
                if(!exists $this_patient_ID_and_patient_diagnosis{$tempid}){
                    # 写入 患者的 临床诊断 信息
                    if ($diagnosis{$num3[$z]} ne "-"){  # "术后患者" 的临床诊断不为 "-"
                        print "L2506|$tempid|patient_diagnosis:$diagnosis{$num3[$z]}\n";
                        $this_patient_ID_and_patient_diagnosis{$tempid} = $diagnosis{$num3[$z]} ;  # 根据 供患信息中 "诊断" 设置
                    }else{
                        print "L2509|$tempid|patient_diagnosis:$diagnosis{$num3[$z]}\n";
                        $this_patient_ID_and_patient_diagnosis{$tempid} = $diagnosis{$num1[$z]} ;  # 根据术前患者的 实验编码 对应 供患信息中 "诊断" 设置
                    }
                }
                if(!exists $this_patient_ID_and_patient_sampleType{$tempid}){
                    my $patient_sampleType = $sample{$num3[$z]};
                    #if ($cells{$num3[$z]} =~ /(\S+)分选/){  # 供患信息中 实验编码 对应的 分选类型 (B细胞分选 / T细胞分选 / NK细胞分选 ...)
                    #    $patient_sampleType = $1;  # 根据 "分选类型" 得到: B细胞/T细胞/NK细胞/粒细胞
                    #}
                    #else{
                    #    $patient_sampleType = $sample{$num3[$z]};  # 否则，将 $patient_sampleType 设置为 "术后患者" 实验编码对应的 "样本类型"
                    #}
                    print "L2523|$tempid|patient_sampleType:$patient_sampleType\n";
                    $this_patient_ID_and_patient_sampleType{$tempid} = $patient_sampleType ;
                }
                if(!exists $this_patient_ID_and_patient_sampleDate{$tempid}){
                    print "L2529|$tempid|patient_sampleDate:$date3{$num3[$z]}\n";
                    $this_patient_ID_and_patient_sampleDate{$tempid} = $date3{$num3[$z]} ;
                }
                if(!exists $this_patient_ID_and_patient_rcvDate{$tempid}){
                    print "L2533|$tempid|patient_rcvDate:$date4{$num3[$z]}\n";
                    $this_patient_ID_and_patient_rcvDate{$tempid} = $date4{$num3[$z]} ;
                }
                if(!exists $this_patient_ID_and_donor_name{$tempid}){
                    print "L2537|$tempid|donor_name:$name{$num2[$z]}\n";
                    $this_patient_ID_and_donor_name{$tempid} = $name{$num2[$z]} ;
                }
                if(!exists $this_patient_ID_and_donor_gender{$tempid}){
                    print "L2541|$tempid|donor_gender:$gender{$num2[$z]}\n";
                    $this_patient_ID_and_donor_gender{$tempid} = $gender{$num2[$z]} ;
                }

                if(!exists $this_patient_ID_and_donor_age{$tempid}){
                    print "L2546|$tempid|donor_age:$age{$num2[$z]}\n";
                    $this_patient_ID_and_donor_age{$tempid} = $age{$num2[$z]} ;
                }

                if(!exists $this_patient_ID_and_donor_relationship{$tempid}){
                    print "L2551|$tempid|dono_relationship:$relation{$num2[$z]}\n";
                    $this_patient_ID_and_donor_relationship{$tempid} = $relation{$num2[$z]} ;
                }

                if(!exists $this_patient_ID_and_hospital{$tempid}){
                    print "L2556|$tempid|hospital:$hospital{$num3[$z]}\n";
                    $this_patient_ID_and_hospital{$tempid} = $hospital{$num3[$z]} ;
                }
                if(!exists $this_patient_ID_and_doctor{$tempid}){
                    print "L2560|$tempid|doctor:$doctor{$num3[$z]}\n";
                    $this_patient_ID_and_doctor{$tempid} = $doctor{$num3[$z]} ;
                }
                # 获取当前报告单 对应的 sampleDetailType
                my $patient_sampleDetailType = "" ;  print "L2575:$cells{$num3[$z]}\n" ;
                if ($cells{$num3[$z]} ne "-"){
                    $patient_sampleDetailType = $sample{$num3[$z]} . "-" . $cells{$num3[$z]} ;
                } else {
                    $patient_sampleDetailType = $sample{$num3[$z]} ;
                }
                print "L2581|$tempid|patient_sampleDetailType:$patient_sampleDetailType\n" ;
                if(!exists $this_patient_ID_sampleDetailType_and_conclusion{$tempid}{$patient_sampleDetailType}){
                    print "L2683|$tempid|$patient_sampleDetailType|conclusion:$conclusion[$z]\n";
                    $this_patient_ID_sampleDetailType_and_conclusion{$tempid}{$patient_sampleDetailType} = $conclusion[$z] ;
                }

                if (!exists $this_patient_ID_and_shuqian_patient_expid{$tempid}) {
                    $this_patient_ID_and_shuqian_patient_expid{$tempid} = $num1[$z] ;
                    print "L2608:$tempid|shuqian_patient_expid|$num1[$z]|$allele{$num1[$z]}{D8S1179}\n" ;
                }

                if (!exists $this_patient_ID_and_shuqian_donor_expid{$tempid}) {
                    $this_patient_ID_and_shuqian_donor_expid{$tempid} = $num2[$z] ;
                    print "L2613:$tempid|shuqian_donor_expid|$num2[$z]|$allele{$num2[$z]}{D8S1179}\n" ;
                }

                # 每种 样本类型的检测结果 分别输出一行
                $hash_this_patient_ID_and_shuqian_patient_genotypes{$tempid} = $allele{$num1[$z]} ;  # 获取 术前患者 的型别结果 : marker => genotype
                $hash_this_patient_ID_and_shuqian_donor_genotypes{$tempid} = $allele{$num2[$z]} ;  # 获取 术前患者 的型别结果 : marker => genotype
                if ($patient_sampleDetailType !~ /\-/) {  # 外周血 / 骨髓血
                    if (!exists $this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_expid{$tempid}) {
                        $this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_expid{$tempid} = $num3[$z] ;
                    }

                    $hash_this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_genotypes{$tempid} = $allele{$num3[$z]}  ; # 获取 术后患者 外周血/骨髓血 样本的型别结果 : marker => genotype
                } elsif ($patient_sampleDetailType =~ /\-T细胞/) {  # T细胞
                    if (!exists $this_patient_ID_and_shuhou_patient_T_cell_expid{$tempid}) {
                        $this_patient_ID_and_shuhou_patient_T_cell_expid{$tempid} = $num3[$z] ;
                    }

                    $hash_this_patient_ID_and_shuhou_patient_T_cell_genotypes{$tempid} = $allele{$num3[$z]}  ; # 获取 术后患者 T细胞 样本的型别结果 : marker => genotype
                } elsif ($patient_sampleDetailType =~ /\-B细胞/) {  # B细胞
                    if (!exists $this_patient_ID_and_shuhou_patient_B_cell_expid{$tempid}) {
                        $this_patient_ID_and_shuhou_patient_B_cell_expid{$tempid} = $num3[$z] ;
                    }

                    $hash_this_patient_ID_and_shuhou_patient_B_cell_genotypes{$tempid} = $allele{$num3[$z]}  ; # 获取 术后患者 B细胞 样本的型别结果 : marker => genotype
                } elsif ($patient_sampleDetailType =~ /\-NK细胞/) {  # NK细胞
                    if (!exists $this_patient_ID_and_shuhou_patient_NK_cell_expid{$tempid}) {
                        $this_patient_ID_and_shuhou_patient_NK_cell_expid{$tempid} = $num3[$z] ;
                    }

                    $hash_this_patient_ID_and_shuhou_patient_NK_cell_genotypes{$tempid} = $allele{$num3[$z]}  ; # 获取 术后患者 NK细胞 样本的型别结果 : marker => genotype
                } elsif ($patient_sampleDetailType =~ /\-粒细胞/) {  # 粒细胞
                    if (!exists $this_patient_ID_and_shuhou_patient_li_cell_expid{$tempid}) {
                        $this_patient_ID_and_shuhou_patient_li_cell_expid{$tempid} = $num3[$z] ;
                    }

                    $hash_this_patient_ID_and_shuhou_patient_li_cell_genotypes{$tempid} = $allele{$num3[$z]}  ; # 获取 术后患者 粒细胞 样本的型别结果 : marker => genotype
                } else {
                    print "L2618: 错误的 patient_sampleDetailType:$patient_sampleDetailType \n请检查!" ;
                }


                # 写入 "备注" 表头
                $worksheet->write('B44',decode('GB2312','备    注'),$format15);
                # 写入 "备注" 的信息
                $worksheet->merge_range('C44:I44', decode('GB2312','1、嵌合状态界定[1]
                    完全嵌合状态(CC): DC≥95%; 混合嵌合状态(MC): 5%≤DC<95%； 微嵌合状态: DC<5%。
                    [1] Outcome of patients with hemoglobinopathies given either cord blood or bone marrow
                    transplantation from an HLA-idebtucak sibling.Blood.2013,122(6):1072-1078.
                    2、本报告用于生物学数据比对、分析，非临床检测报告。'), $format19);

                # 写入 "检测者" "复核者" "检测日期" "报告日期"
                $worksheet->merge_range('B46:C46', decode('GB2312','检  测  者'), $format7);
                $worksheet->merge_range('B47:C47', decode('GB2312','复  核  者'), $format7);
                $worksheet->merge_range('D46:E46', decode('GB2312',''), $format7);
                $worksheet->merge_range('D47:E47', decode('GB2312',''), $format7);
                $worksheet->merge_range('F46:G46', decode('GB2312','检测日期'), $format7);
                $worksheet->merge_range('F47:G47', decode('GB2312','报告日期'), $format7);
                $worksheet->merge_range('H46:I46', decode('GB2312',DateUnify($date1{$num3[$z]})), $format9);  # 将 "术后患者" 实验编码 对应的 "收样日期" 设置为 供患信息中 实验编码 对应的 生产时间  （用生产日期 设置 收样日期 ？）
                $worksheet->merge_range('H47:I47', decode('GB2312',sprintf("%d-%02d-%02d",$year,$mon,$mday)), $format9);  # 设置 "报告日期" 为报告软件出报告的 当天

                # 在 "报告" 表单 相应位置，插入 "检测者" "复核者" "盖章" 图片
                if (-e "pic/检测者.png"){
                     $worksheet->insert_image('D46', "pic/检测者.png", 5, 5, 1.1, 1.1);
                }
                if (-e "pic/复核者.png"){
                     $worksheet->insert_image('D47', "pic/复核者.png", 5, 8, 0.8, 0.8);
                }
                if (-e "pic/盖章.png"){
                     $worksheet->insert_image('H46', "pic/盖章.png", 3, -30, 0.998, 0.977);
                }

                #姓名        医院        样本类型        样本编号        报告编号        嵌合率
                # 如果 样本的嵌合率 == 0
                # 重写写入 嵌合率汇总 的结果文件 "嵌合率汇总-%4d%02d%02d.txt"
                if ($count_avg[$z] == 0){
                        printf SUM "%s\t%s\t%s\t%s\t%s\t%f%s\t%d\tNA\tNA\n", $name{$num3[$z]}, $hospital{$num3[$z]}, $sample{$num3[$z]}, $number{$num3[$z]}, $rptnum{$num3[$z]}, $count_avg[$z],"%",$count_n[$z];
                }
                else{
                        #姓名\t医院\t样本类型\t样本编号\t报告编号\t嵌合率\t有效位点\tSD\tCV
                        printf SUM "%s\t%s\t%s\t%s\t%s\t%f%s\t%d\t%.2f%s\t%.2f%s\n", $name{$num3[$z]}, $hospital{$num3[$z]}, $sample{$num3[$z]}, $number{$num3[$z]}, $rptnum{$num3[$z]}, $count_avg[$z],"%",$count_n[$z], $SD[$z]*100,"%", $SD[$z]/$count_avg[$z]*10000,"%";
                }
                # 如果 报告单编号 对应的 实验编码 不为3 或者 $sheet[$z] 对应的 "患者姓名"  的第二份及以后的报告，不生成嵌合曲线
                print "L2610:$z|$TCAID|$exp_num{$TCAID}|$sheet_name{$sheet[$z]}\n" ;
                if ($exp_num{$TCAID} != 3 or $sheet_name{$sheet[$z]} > 1){
                        $workbook -> close();
                        $RptBox -> Append("报告生成成功！无嵌合曲线\r\n");  # 更新 "生成报告" 部分的文本框中显示的提示信息
                        next;
                }
                $RptBox -> Append("报告生成成功！");  # 更新 "生成报告" 部分的文本框中显示的提示信息

                #########################################  2020.05.24.14.32 ##########################################################
                ######################################### 输出 "temp" 表单 ##########################################################
                $tempid = $identity{$TCAID};  # 获取 报告单编号 对应的 患者编码 # %identity 存储 报告单编号 <=> 患者编码(HUN001胡琳)
                my $i;
                my $j = 1;
                my $Chart_Marker_Num = 0;
                my %Graphic_SampleID;
                my %Graphic_Chimerism;
                my %Types;
                my @date_seq;
                push @date_seq, 0;
                if(exists $Chimerism{$tempid}){  #  判断 患者编码 对应的嵌合结果 在 %Chimerism 里是否存在 # %Chimerism 将该 患者编码 对应的嵌合结果（有效位点嵌合率的平均值）存入$Chimerism{$tempid}
                        foreach $i(0 .. $#{$Chimerism{$tempid}}){  # 遍历 患者编码 对应的几份报告的嵌合率
                                my $Chmrsm = $Chimerism{$tempid}[$i];  # "患者编码" 对应的第几份报告的 嵌合率
                                $Chmrsm =~ s/%//;  # 嵌合率是 百分比格式
                                $Chmrsm = sprintf ("%.2f", $Chmrsm);  # 将 百分比格式 转换为 小数（保留小数点后2位）
                                my $Smplid = $SampleID{$tempid}[$i];   # 实验编码 # 数组里每一个元素都是 hash表，用于存储每个 患者编码 对应的每次检测的 实验编码
                                my $SmpType = $sampleType{$Smplid};  # 供患信息中 实验编码 对应的 分选类型 / 样本类型
                                next unless $SmpType;  # 如果样本类型为空，则跳到下一份报告单
                                next if $SmpType eq "-";  # 如果样本类型为 "-"，则跳到下一份报告单
                                my $rptDate = DateUnify($ReportDate{$tempid}[$i]);  # 获取 患者编码 对应的 第i份 报告的 报告日期
                                my $rcvDate = DateUnify($receiveDate{$Smplid});  # 获取 患者编码 对应的 第i份 报告的 收样日期
                                my $smplDate = DateUnify($sampleDate{$Smplid});  # 获取 患者编码 对应的 第i份 报告的 采样日期
                                my $tmpDate;
                                if ($smplDate ne '不详' && $smplDate ne '-'){  # 采样日期 不为 "不详" 或 "-"
                                        $tmpDate = $smplDate;
                                }elsif ($rcvDate ne '不详' && $rcvDate ne '-'){  # 收样日期 不为 "不详" 或 "-"
                                        $tmpDate = $rcvDate;
                                }elsif($rptDate ne '不详'){  # 报告日期 不为 "不详"
                                        $tmpDate = $rptDate;
                                }else{  # 采样日期 / 收样日期 / 报告日期 都是 "不详" 或者 "-"
                                        $tmpDate = sprintf "%s%d%s", "术后", $j, "次";
                                }

                                # print $Smplid,"|", $rptDate,"|",$rcvDate,"|",$smplDate,"|",$tmpDate,"\n";

                                $Graphic_Chimerism{$tmpDate}{$SmpType} = $Chmrsm;  # 存储 样本 某个时间（采样日期 / 收样日期 / 报告日期）对应的 某个类型的样本 对应的 嵌合率结果
                                $Graphic_SampleID{$tmpDate}{$SmpType} = $Smplid;  # 存储 样本 某个时间（采样日期 / 收养日期 / 报告日期）对应的 某个类型的样本 对应的 实验编码
                                $Types{$SmpType} ++;  print "L2649:$tempid|$TCAID|" . $SmpType . ":" . $Types{$SmpType} . "\n" ;# 样本类型 / 分选类型 个数递增 (这里 $Types{$SmpType} 没有判断是否存在，也没有设置初始化的值？ 默认初始化的值为 0 ? 测试一下看看)
                                if ($date_seq[-1] ne $tmpDate || $tmpDate =~ /术后/){  # 将样本的 （采样日期 / 收样日期 / 报告日期）存入 @date_seq
                                        push @date_seq, $tmpDate;
                                        $j ++;
                                }
                                $Chart_Marker_Num ++;  # 记录总的报告数？
                        }

                        ############################ 写入 temp 表单 #######################################
                        shift @date_seq;
                        my $headings;
                        push @{$headings}, decode('GB2312', '时间');
                        my $write_data;
                        foreach (@date_seq){
                                push @{${$write_data}[0]}, decode('GB2312', $_);
                        }
                        $i = 1;
                        foreach (keys %Types){
                                push @{$headings}, decode('GB2312', $_);
                                foreach my $tmp(@date_seq){
                                        if (exists $Graphic_Chimerism{$tmp}{$_}){
                                                push @{${$write_data}[$i]}, decode('GB2312', $Graphic_Chimerism{$tmp}{$_});
                                        }else{
                                                push @{${$write_data}[$i]}, undef;
                                        }
                                }
                                $i ++;
                        }

                        $graphic_temp -> write('A1', $headings);  # 写 表头 列
                        $graphic_temp -> write('A2', $write_data);  # 写入 每次检测的嵌合率结果


                        ############################ 输出 "嵌合曲线" 表单 #######################################
                        $graphic->hide_gridlines();  # 隐藏 网格线
                        $graphic->keep_leading_zeros();  # 保留 开头的0
                        # 设置各列的宽度
                        $graphic->set_column(0,0,4.24);
                        $graphic->set_column(1,6,13.75);
                        $graphic->set_column(7,7,4.24);
                        # 设置各行的高度
                        my @rows = (75,   3.6, 3.6 , 3.6 , 19 ,  18 ,  15.6, 15.6,
                                                15.6, 18.6,        25.8, 18.6, 19.2, 24.6, 19.8, 16.2,
                                                16.2, 16.2, 16.2, 16.2,        16.2, 16.2, 16.2, 24, #24 from 32 to 24
                                                18.75,16.25,16.25,16.25,33.0, 16.25, 16.25);#last from 16.25 to 17
                        for my $i(0 .. $#rows){
                                $graphic->set_row($i, $rows[$i]);
                        }

                        # 设置 "嵌合曲线" 表单 下面 每次检测结果表格的行高
                        foreach $i(1 .. $Chart_Marker_Num){
                                $graphic->set_row($i+30, 13.5);
                        }

                        # 设置 左右和上 边距
                        $graphic->set_margin_left(0.394);
                        $graphic->set_margin_right(0.394);
                        #$graphic->set_margin_top(0.2);

                        # 设置 页脚
                        # $graphic->set_footer($footer);  # $footer: L2418 - L2421

                        # B1 插入 logo
                        $graphic->insert_image('B1', "pic/logo.png", 10, 10, 0.15, 0.15);

                        # 写入 "嵌合曲线" 表头
                        $graphic->merge_range('B1:G1', decode('GB2312','嵌合曲线'), $format1);
                        #$graphic->merge_range('B2:G2', decode('GB2312','地址：上海市浦东新区紫萍路908弄21号（上海国际医学园区）          邮编：201318'), $format2);
                        # 插入一个空行
                        $graphic->merge_range('B4:G4', decode('GB2312',''), $format2);

                        # 写入 "患者姓名" 表头
                        $graphic->write('B5',decode('GB2312','患者姓名'), $Gfmt1);
                        # 写入 "患者姓名" 内容 $name[$num3[$z]]
                        $graphic->write('C5',decode('GB2312',$name{$num3[$z]}), $Gfmt2);
                        # 写入 "样本编号" 表头 及 内容 （实际为 供患信息中的 "实验编码"）
                        $graphic->write('F5',decode('GB2312','样本编号：'), $Gfmt1);
                        $graphic->write('G5',decode('GB2312',$number{$num3[$z]}), $Gfmt1);

                        # 插入嵌合曲线图
                        my $chart = $workbook->add_chart(type => 'line', embedded => 1 );
                        my $row_max = $#{${$write_data}[0]}+1; # 获取 "temp" 表单的总行数， +1 是因为 add_series 里是以1为起始的
                        my $col_max = $#{$write_data};  # 获取 "temp" 表单的总列数
                        for my $i(1..$col_max){  # 遍历加入  每个 样本类型
                                my $formula = sprintf "=temp!\$%s1", chr($i+65);  # 格式后续测试时关注
                                $chart->add_series(  # 根据 "temp" 表单，选取画图的数据区域
                                        categories => ['temp', 1,$row_max, 0 , 0],  # 选取 "temp" 的第一列作为 categories
                                        values     => ['temp', 1, $row_max, $i, $i],  # 选择 "temp" 的第i+1列 作为 values (对应每一种样本类型)
                                        name_formula => $formula,
                                        # name       => decode('GB2312',${$headings}[$i]),
                                                marker   => {  # 设置每个 series (样本类型) 的符号
                                                        type    => 'automatic',
                                                        size    => 5,
                                                },
                                );
                        }

                        #
                        $chart->set_chartarea(  # is used to set the properties of the chart area
                                color => 'white',
                                line_color => 'black',
                                line_weight => 3,
                        );

                        $chart->set_plotarea(  # is used to set properties of the plot area of a chart.
                                color => 'white',

                        );

                        $chart->set_y_axis(  # is used to set properties of the Y axis.
                                name => decode('GB2312','嵌合率(%)'),
                                min  => 0,
                                max  => 100,
                                major_unit => 20,
                        );

                        # 设置 图例的位置(bottom) / 整个图表的宽高(in pixels)
                        $chart->set_legend( position => 'bottom' );
                        $chart->set_size( width => 607, height => 400 );
                        # 将图表插入 "嵌合曲线" 的 "B7"
                        $graphic->insert_chart('B7', $chart);

                        # "嵌合曲线" 图表下面 提示信息 部分的内容
                        # 写入 "TCA定期检测流程 .... 等说明性 内容"
                        $graphic->merge_range('B25:G25', decode('GB2312','TCA定期检测流程'), $Gfmt5);
                        $graphic->merge_range('B26:G26', decode('GB2312','基线检测：术前同时对供、受者进行检测、也可以在术后首次追踪检测时进行'), $Gfmt6);
                        $graphic->merge_range('B27:G27', decode('GB2312','追踪检测：建议在术后2周进行第一次TCA，第4周进行第二次检测；'),$Gfmt6);
                        $graphic->merge_range('B28:G28', decode('GB2312','        术后6个月内，每月检测一次；6个月之后，每2个月检测一次，直至嵌合率稳定'), $Gfmt6);
                        # 插入建议的检测时间轴 "pic/comment.bmp"
                        $graphic->insert_image('B29', "pic/comment.bmp", 5, 5);
                        $graphic->merge_range('B30:G30', decode('GB2312','温馨提示：一旦术后免疫治疗方案调整，在调整后2周需要重新启动检测'), $Gfmt7);

                        # 写入 "嵌合曲线" 最下部分 "每次检测的 采样日期/检测日期/嵌合率(%)/样本编号(实际为实验编码)/样本类型" 汇总表
                        # 写入表头
                        $graphic->write('B32', decode('GB2312','检测次数'), $Gfmt3);
                        $graphic->write('C32', decode('GB2312','采样日期'), $Gfmt3);
                        $graphic->write('D32', decode('GB2312','检测日期'), $Gfmt3);
                        $graphic->write('E32', decode('GB2312','嵌合率(%)'),   $Gfmt3);
                        $graphic->write('F32', decode('GB2312','样本编号'), $Gfmt3);
                        $graphic->write('G32', decode('GB2312','样本类型'), $Gfmt3);

                        my $i = 1;
                        my $j = 1;
                        for my $tmpDate(@date_seq){  # 遍历每个采样日期
                                for my $SmpType(keys %Types){  # 遍历每个采样日期的每种 样本类型
                                        my $Smplid = $Graphic_SampleID{$tmpDate}{$SmpType};  # 某个采样日期的某个样本类型 对应的 "实验编码"
                                        my $Chmrsm = $Graphic_Chimerism{$tmpDate}{$SmpType};  # 某个采样日期的某个样本类型 对应的 "嵌合率"
                                        next unless $Smplid;  # 跳过 "实验编码" 为空的行
                                        my $rcvDate = $receiveDate{$Smplid};  # 获取 "实验编码" 对应的 "收样日期"
                                        my $smplDate = $sampleDate{$Smplid};  # 获取 "实验编码" 对应的 "采样日期"

                                        $graphic->write($j+31, 1, $i, $Gfmt4);  # 写入 检测次数
                                        $graphic->write($j+31, 2, decode('GB2312',$smplDate), $Gfmt4);  # 写入 "采样日期"
                                        $graphic->write($j+31, 3, decode('GB2312',$rcvDate), $Gfmt4);  # 写入 "收样日期" （对应表格中的 "检测日期"）
                                        $graphic->write($j+31, 4, sprintf("%.2f",$Chmrsm), $Gfmt8);  # 写入 "嵌合率" 结果，按百分比个数，保留小数点后2位有效数字
                                        $graphic->write($j+31, 5, $Smplid, $Gfmt4);  # 写入 "实验编码" （对应表格中的 "样本编号"）
                                        $graphic->write($j+31, 6, decode('GB2312',$SmpType), $Gfmt9);  # 写入 "样本类型"
                                        $j ++;
                                        if (($j-11)%54 == 0){  # 当检测次数超过5次时，重新写一次表头
                                                $graphic->write($j+31, 1, decode('GB2312','检测次数'), $Gfmt3);
                                                $graphic->write($j+31, 2, decode('GB2312','采样日期'), $Gfmt3);
                                                $graphic->write($j+31, 3, decode('GB2312','检测日期'), $Gfmt3);
                                                $graphic->write($j+31, 4, decode('GB2312','嵌合率(%)'),   $Gfmt3);
                                                $graphic->write($j+31, 5, decode('GB2312','样本编号'), $Gfmt3);
                                                $graphic->write($j+31, 6, decode('GB2312','样本类型'), $Gfmt3);
                                                $j ++;
                                        }
                                }
                                $i ++;
                        }
                }else{

                }

                # excel 文件写完，关闭文件
                $workbook->close();
                $RptBox -> Append("嵌合曲线生成成功！\r\n");  # 更新 "生成报告" 部分的文本框中显示的提示信息
        }

        # 输出 状态 提示信息 ($sb))
        $sb->Move( 0, ($main->ScaleHeight() - $sb->Height()) );
        $sb->Resize( $main->ScaleWidth(), $sb->Height() );
        if ($success){
                $sb->Text("输出完成");
        }else{
                $sb->Text("输出完成（有错误）");
        }

        $RUNwindow -> Hide();
        if ($success){
                $error =  "输出保存成功！\n";
                Win32::MsgBox $error, 0, "成功！";
        }else{
                $error =  "输出保存成功，但发生了错误！\n";
                Win32::MsgBox $error, 0, "注意！";
        }

        close SUM;

        #############################  以 患者编码 为单位，按JRK的报告模板，输出每个患者 每次检测的多份样本的检测结果 ######################
        $sb->Move( 0, ($main->ScaleHeight() - $sb->Height()) );
        $sb->Resize( $main->ScaleWidth(), $sb->Height() );
        $sb->Text("开始生成每个患者多份检测样本的合并报告，正在生成文件...");  # 状态条，输出提示信息
        print "L2766:开始生成每个患者多份检测样本的合并报告，正在生成文件...\n" ;
        # 遍历所有 患者编码
        my $tt = 0 ;
        for my $tempid (keys %Chimerism){
            print "L2771:$tt\n" ;
            if($tt >= 5){ print "Get out .. \n"; last; }
            print "L2769:", $tempid,"|", $#{$Chimerism{$_}}+1,"\n";

            # 定义 报告单 输出文件
            ####### 需要获取 患者 本轮检测的报告单编号 ######
            # my $ana_date = sprintf "%4d%02d%02d.txt",$year, $mon, $mday; ; print "ana_date: $ana_date\n";
            # $Output_Dir\\$name{$num3[$z]}-STR检测报告-$hosptl_num{$num3[$z]}-$ana_date
            my $Output_rpt_str = sprintf "%s\\%s-STR检测报告-%s-%s\\%s-STR检测报告-%s-%s.xlsx", $Output_Dir, $this_patient_ID_and_patient_name{$tempid}, $this_patient_ID_and_hospital{$tempid}, $ana_date,$this_patient_ID_and_patient_name{$tempid}, $this_patient_ID_and_hospital{$tempid}, $ana_date;
            print "L2778:$Output_rpt_str\n" ;
            my $workbook;
            # 打开 报告输出文件，准备写报告
            unless ($workbook = Excel::Writer::XLSX->new($Output_rpt_str)){
                $error = $Output_rpt_str."无法保存！";
                Win32::MsgBox $error, 0, "错误！";
                $success = 0;
                $RptBox -> Append($Output_rpt_str."打开失败！跳过\r\n");
                next;
            }

            my ($countsheet, $graphic, $worksheet, $graphic_temp);
            $worksheet  = $workbook ->add_worksheet(decode('GB2312',"报告"));  # 增加 "报告" 表单
            $graphic = $workbook ->add_worksheet(decode('GB2312',"嵌合曲线"));  # 增加 "嵌合曲线" 表单
            # $countsheet = $workbook->add_worksheet(decode('GB2312',"计算"));  # 汇总的报告中，暂时不输出 "计算" 表单 （信息太多，可到单个类型报告中核对）
            $graphic_temp = $workbook->add_worksheet('temp');  # 增加 "temp" 表单

            ##################################################################################################################################
            ########################################### 输出 "报告" 表单 #######################################################################
            # 定义 "报告" 表单的各种格式
            $worksheet->hide_gridlines();
            $worksheet->keep_leading_zeros();
            # 设置 列的宽度、行的高度
            # $worksheet->set_column(0,0,0.5);
            $worksheet->set_column(0,0,12);
            $worksheet->set_column(1,1,11);
            $worksheet->set_column(2,3,11);
            $worksheet->set_column(4,5,11);
            $worksheet->set_column(6,7,11);
            # $worksheet->set_column(8,8,10);
            my @rows = (45,1,1,28,18,14,14,14,14,14,14,14,14,14,5,18,18,18,18,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,14,35,10,10,10,10);
            for my $i (0 .. $#rows){$worksheet->set_row($i, $rows[$i]);}

            # 设置页面的左右及上边距
            $worksheet->set_margin_left(0.3);
            $worksheet->set_margin_right(0.3);
            $worksheet->set_margin_top(0.25);
            $worksheet->set_margin_bottom(0.25);

            ###################  定义excel文件各部分的格式  #######################################################
            my $format1  = $workbook->add_format(size => 16, bold => 1, align => 'center',                      font => decode('GB2312','宋体'), color => '#00B0F0', 'bottom' => 2); # "广州君瑞康生物科技有限公司"
            my $format2  = $workbook->add_format(size => 11);  # 不画 框线
            my $format3  = $workbook->add_format(size => 16, bold => 1, align => 'center', valign => 'vcenter', font => decode('GB2312','宋体')); # "移植后嵌合体状态分析报告"
            my $format4  = $workbook->add_format(size => 12,            align => 'right',  valign => 'vcenter', font => decode('GB2312','宋体')); # "报告单编号"
            my $format5  = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 患者 及 供者信息表 单元格 格式
            my $format6  = $workbook->add_format(size => 11, bold => 1, align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1,                 ); # "检测结果:"
            my $format7  = $workbook->add_format(size => 11, bold => 1, align => 'left', valign => 'vcenter', font => decode('GB2312','宋体'),                             ); # 检测结果 内容
            my $format8  = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # "基因座"
            my $format9  = $workbook->add_format(size => 11,            align => 'center', valign => 'bottom', font => decode('GB2312','宋体'), 'top' => 1,                'left' => 1, 'right' => 1); # "样本编号"
            my $format10 = $workbook->add_format(size => 7,             align => 'center', valign => 'top', font => 'Times New Roman',                   'bottom' => 1, 'left' => 1, 'right' => 1); # 样本编号 内容
            my $format11 = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1,  'bottom' => 1, 'left' => 1, 'right' => 1); # "患者移植前" 和 "供者"
            my $format12 = $workbook->add_format(size => 11,            align => 'center', valign => 'bottom', font => decode('GB2312','宋体'), 'top' => 1,                 'left' => 1, 'right' => 1); # "患者移植后"
            my $format13 = $workbook->add_format(size => 11,            align => 'center', valign => 'top', font => decode('GB2312','宋体'),              'bottom' => 1, 'left' => 1, 'right' => 1); # "(骨髓血)/(外周血)" / "(T细胞)" / "(B细胞)" / "(NK细胞)" / "(粒细胞)"
            my $format14 = $workbook->add_format(size => 10,            align => 'center', valign => 'vcenter', font => 'Times New Roman',       'top' => 1,  'bottom' => 1, 'left' => 1, 'right' => 1);  # marker名称
            my $format15 = $workbook->add_format(size => 10,            align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'),  'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # marker 的型别
            my $format16 = $workbook->add_format(size => 11,            align => 'right', valign => 'vcenter', font => decode('GB2312','宋体'));  # "检测者" | "复核者" | "报告日期" 及其 内容
            my $format21 = $workbook->add_format(size => 11,            align => 'left', valign => 'vcenter', font => decode('GB2312','宋体'));  #  报告日期 内容
            my $format17  = $workbook->add_format(size => 11,           align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # "备注"
            my $format18 = $workbook->add_format(size => 8, valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'left' => 1, 'right' => 1, 'text_wrap' => 1); #备注的 第一行 内容
            my $format19 = $workbook->add_format(size => 8, valign => 'vcenter', font => decode('GB2312','宋体'),             'left' => 1, 'right' => 1, 'text_wrap' => 1); #备注的 第二-三行 内容
            my $format20 = $workbook->add_format(size => 8, valign => 'top',     font => decode('GB2312','宋体'), 'bottom' => 1, 'left' => 1, 'right' => 1, 'text_wrap' => 1); #备注的 第四行 内容
            # my $format22 = $workbook->add_format(size => 11,            align => 'l', valign => 'ctr', font => decode('GB2312','宋体'));  #  报告日期 内容

            # 设置表单的 页脚
            #my $footer = '&L'.decode('GB2312','公司地址：广州市黄埔区新瑞路6号二栋2层A205房、A206房')."\n".
            #             #'&R'.decode('GB2312','XXXX-X-X-0021-1.0')."\n".
                         '&L'.decode('GB2312','咨询电话：020-62313880').
            #             #'&R'.decode('GB2312',$this_patient_ID_and_patient_name{$tempid}.'，第').'&P'.decode('GB2312','页/共').'&N'.decode('GB2312','页');
                          '&C&9'.decode('GB2312','&P'.'/'.'&N');
            my $footer = '&C&9'.decode('GB2312','&P'.'/'.'&N');
            $worksheet->set_footer($footer);

            # 在第一行第二列插入 公司logo (pic/logo.png)
            $worksheet->insert_image('A1', "pic/logo.png", 1, 16, 0.17623, 0.176);

            # 写入 各项信息
            ################## 报告头 部分 #########################
            # $worksheet->merge_range('A1:H1', decode('GB2312','广州君瑞康生物科技有限公司'), $format1);
            my $bold = $workbook->add_format(size => 12, bold => 1 , font => decode('GB2312','宋体'));
            my $normal = $workbook->add_format(size => 11, font => decode('GB2312','宋体'));
            my $fmt_telephone = $workbook->add_format(size => 10, font => decode('GB2312','宋体'));
            my $fmt_align = $workbook->add_format( align => 'right', valign => 'top', 'bottom' => 2);
            $fmt_align->set_text_wrap();
            $worksheet->merge_range_type( 'rich_string', 'A1:H1', $bold, decode('GB2312',"广州君瑞康生物科技有限公司"), "\n", $normal, decode('GB2312',"广州市黄埔区新瑞路6号二栋A205"), "\n", $fmt_telephone, decode('GB2312',"020-62313880"), "\n", $fmt_align);
            $worksheet->merge_range('A2:H2', decode('GB2312',''), $format2);  # 留 空行
            $worksheet->merge_range('A3:H3', decode('GB2312',''), $format2);  # 留 空行
            $worksheet->merge_range('A4:H4', decode('GB2312','移植后嵌合体状态分析报告'), $format3);  #

            ################## 报告单编号 部分 #########################
            $worksheet->merge_range('A5:H5', decode('GB2312','报告编号：'.$this_patient_ID_and_report_id{$tempid}),$format4);

            ################## 供患信息 部分 #########################
            $worksheet->merge_range('A6:A7', decode('GB2312','患者姓名'), $format5);
            $worksheet->merge_range('B6:B7', decode('GB2312',$this_patient_ID_and_patient_name{$tempid}), $format5);
            $worksheet->write('C06',decode('GB2312','性别'),$format5);
            $worksheet->write('D06',decode('GB2312',$this_patient_ID_and_patient_gender{$tempid}),$format5);
            $worksheet->write('E06',decode('GB2312','年龄'),$format5);
            $worksheet->write('F06',decode('GB2312',$this_patient_ID_and_patient_age{$tempid}),$format5);
            $worksheet->write('G06',decode('GB2312','采样日期'),$format5);
            $worksheet->write('H06',decode('GB2312',$this_patient_ID_and_patient_sampleDate{$tempid}),$format5);

            $worksheet->write('C07',decode('GB2312','临床描述'),$format5);
            $worksheet->write('D07',decode('GB2312',$this_patient_ID_and_patient_diagnosis{$tempid}),$format5);
            $worksheet->write('E07',decode('GB2312','样本类型'),$format5);
            $worksheet->write('F07',decode('GB2312',$this_patient_ID_and_patient_sampleType{$tempid}),$format5);
            $worksheet->write('G07',decode('GB2312','接样日期'),$format5);
            $worksheet->write('H07',decode('GB2312',$this_patient_ID_and_patient_rcvDate{$tempid}),$format5);

            $worksheet->write('A8', decode('GB2312','供者姓名'), $format5);
            $worksheet->write('B8', decode('GB2312',$this_patient_ID_and_donor_name{$tempid}), $format5);
            $worksheet->write('C08',decode('GB2312','性别'),$format5);
            $worksheet->write('D08',decode('GB2312',$this_patient_ID_and_donor_gender{$tempid}),$format5);
            $worksheet->write('E08',decode('GB2312','年龄'),$format5);
            $worksheet->write('F08',decode('GB2312',$this_patient_ID_and_donor_age{$tempid}),$format5);
            $worksheet->write('G08',decode('GB2312','供患关系'),$format5);
            $worksheet->write('H08',decode('GB2312',$this_patient_ID_and_donor_relationship{$tempid}),$format5);

            $worksheet->write('A09',decode('GB2312','送检单位'),$format5);
            $worksheet->merge_range('B9:D9', decode('GB2312',$this_patient_ID_and_hospital{$tempid}), $format5);
            $worksheet->write('E09',decode('GB2312','送检专家'),$format5);
            $worksheet->merge_range('F9:H9', decode('GB2312',$this_patient_ID_and_doctor{$tempid}), $format5);

            ################## 检测结果 部分 #########################
            # $worksheet->merge_range('A10:H10',decode('GB2312',' 检测结果：'),$format6);
            $worksheet->write('A10',decode('GB2312','  检测结果：'),$format6);
            # 每种 样本类型的检测结果 分别输出一行
            my $hash_this_patient_conclusions = $this_patient_ID_sampleDetailType_and_conclusion{$tempid} ;
            my @this_patient_sampleDetailTypes = keys %$hash_this_patient_conclusions ;
            print "L3034:". @this_patient_sampleDetailTypes ."\n" ;
            # 按照 "外周血/骨髓血" "T细胞" "B细胞" "NK细胞" "粒细胞" 的顺序，输出检测结果
            my $conclusion_row_index = 0 ;
            for my $q (0..$#this_patient_sampleDetailTypes){
                if ($this_patient_sampleDetailTypes[$q] !~ /\-/){  # 不是分选类型
                    print "L3038:$conclusion_row_index|$this_patient_sampleDetailTypes[$q]|$hash_this_patient_conclusions->{$this_patient_sampleDetailTypes[$q]}\n";
                    $worksheet->merge_range($conclusion_row_index+10,0,$conclusion_row_index+10,7,decode('GB2312','    '.$hash_this_patient_conclusions->{$this_patient_sampleDetailTypes[$q]}), $format7);  # 写入每一类样本的结论
                    $conclusion_row_index ++ ;
                }
            }
            for my $q (0..$#this_patient_sampleDetailTypes){
                if ($this_patient_sampleDetailTypes[$q] =~ /-T细胞/){  # T细胞分选类型
                    print "L3038:$conclusion_row_index|$this_patient_sampleDetailTypes[$q]|$hash_this_patient_conclusions->{$this_patient_sampleDetailTypes[$q]}\n";
                    $worksheet->merge_range($conclusion_row_index+10,0,$conclusion_row_index+10,7,decode('GB2312','    '.$hash_this_patient_conclusions->{$this_patient_sampleDetailTypes[$q]}), $format7);  # 写入每一类样本的结论
                    $conclusion_row_index ++ ;
                }
            }
            for my $q (0..$#this_patient_sampleDetailTypes){  # 写 "B细胞分选结果"
                if ($this_patient_sampleDetailTypes[$q] =~ /-B细胞/){  # B细胞分选类型
                    print "L3038:$conclusion_row_index|$this_patient_sampleDetailTypes[$q]|$hash_this_patient_conclusions->{$this_patient_sampleDetailTypes[$q]}\n";
                    $worksheet->merge_range($conclusion_row_index+10,0,$conclusion_row_index+10,7,decode('GB2312','    '.$hash_this_patient_conclusions->{$this_patient_sampleDetailTypes[$q]}), $format7);  # 写入每一类样本的结论
                    $conclusion_row_index ++ ;
                }
            }
            for my $q (0..$#this_patient_sampleDetailTypes){  # 写 "NK细胞分选结果"
                if ($this_patient_sampleDetailTypes[$q] =~ /-NK细胞/){  # NK细胞分选类型
                    print "L3038:$conclusion_row_index|$this_patient_sampleDetailTypes[$q]|$hash_this_patient_conclusions->{$this_patient_sampleDetailTypes[$q]}\n";
                    $worksheet->merge_range($conclusion_row_index+10,0,$conclusion_row_index+10,7,decode('GB2312','    '.$hash_this_patient_conclusions->{$this_patient_sampleDetailTypes[$q]}), $format7);  # 写入每一类样本的结论
                    $conclusion_row_index ++ ;
                }
            }
            for my $q (0..$#this_patient_sampleDetailTypes){  # 写 "粒细胞分选结果"
                if ($this_patient_sampleDetailTypes[$q] =~ /-粒细胞/){  # 粒细胞分选类型
                    print "L3038:$conclusion_row_index|$this_patient_sampleDetailTypes[$q]|$hash_this_patient_conclusions->{$this_patient_sampleDetailTypes[$q]}\n";
                    $worksheet->merge_range($conclusion_row_index+10,0,$conclusion_row_index+10,7,decode('GB2312','    '.$hash_this_patient_conclusions->{$this_patient_sampleDetailTypes[$q]}), $format7);  # 写入每一类样本的结论
                    $conclusion_row_index ++;
                }
            }

            # 本次检测所用 样本类型 的型别汇总表
            $worksheet->merge_range('A15:H15', decode('GB2312',''), $format2);
            $worksheet->merge_range('A16:A19', decode('GB2312','基因座'), $format8);
            $worksheet->write('B16', decode('GB2312','样本编号'), $format9);
            $worksheet->write('B17', decode('GB2312',$this_patient_ID_and_shuqian_patient_expid{$tempid}), $format10);
            $worksheet->merge_range('B18:B19', decode('GB2312',"患者移植前"), $format11);
            $worksheet->write('C16', decode('GB2312','样本编号'), $format9);
            $worksheet->write('C17', decode('GB2312',$this_patient_ID_and_shuqian_donor_expid{$tempid}), $format10);
            $worksheet->merge_range('C18:C19', decode('GB2312',"供者"), $format11);
            # 判断样本类型为 外周血 or 骨髓 ？

            my $sample_type = $this_patient_ID_and_patient_sampleType{$tempid}  ;
            $worksheet->write('D16', decode('GB2312','样本编号'), $format9);
            $worksheet->write('D17', decode('GB2312',$this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_expid{$tempid}), $format10);
            $worksheet->write('D18', decode('GB2312','患者移植后'), $format12);
            if ($sample_type eq "外周血" || $sample_type eq "全血"){
                $worksheet->write('D19', decode('GB2312','(外周血)'), $format13);
            } elsif ($sample_type eq "骨髓" || $sample_type eq "骨髓血" ){
                $worksheet->write('D19', decode('GB2312','(骨髓血)'), $format13);
            } else {
                $worksheet->write('D19', decode('GB2312','(其他)'), $format13);
            }

            # 术后患者 T细胞
            $worksheet->write('E16', decode('GB2312','样本编号'), $format9);
            $worksheet->write('E17', decode('GB2312', $this_patient_ID_and_shuhou_patient_T_cell_expid{$tempid}), $format10);
            $worksheet->write('E18', decode('GB2312','患者移植后'), $format12);
            $worksheet->write('E19', decode('GB2312','(T细胞)'), $format13);
            # 术后患者 B细胞
            $worksheet->write('F16', decode('GB2312','样本编号'), $format9);
            $worksheet->write('F17', decode('GB2312',$this_patient_ID_and_shuhou_patient_B_cell_expid{$tempid}), $format10);
            $worksheet->write('F18', decode('GB2312','患者移植后'), $format12);
            $worksheet->write('F19', decode('GB2312','(B细胞)'), $format13);
            # 术后患者 NK细胞
            $worksheet->write('G16', decode('GB2312','样本编号'), $format9);
            $worksheet->write('G17', decode('GB2312',$this_patient_ID_and_shuhou_patient_NK_cell_expid{$tempid}), $format10);
            $worksheet->write('G18', decode('GB2312','患者移植后'), $format12);
            $worksheet->write('G19', decode('GB2312','(NK细胞)'), $format13);
            # 术后患者 粒细胞
            $worksheet->write('H16', decode('GB2312','样本编号'), $format9);
            $worksheet->write('H17', decode('GB2312', $this_patient_ID_and_shuhou_patient_li_cell_expid{$tempid}), $format10);
            $worksheet->write('H18', decode('GB2312','患者移植后'), $format12);
            $worksheet->write('H19', decode('GB2312','(粒细胞)'), $format13);

            # 在第一列 遍历写入25个marker名字
            for my $q (0..$#markers_jrk){
                if ($markers_jrk[$q] =~ /vWA/){$markers_jrk[$q] =~ s/vWA/VWA/;}  # 将marker "vWA" 替换为 "VMA"
                if ($markers_jrk[$q] =~ /AMEL/){$markers_jrk[$q] =~ s/AMEL/Amel/;}  # 将marker "AMEL" 替换为 "Amel"
                $worksheet->write($q+19,0,$markers_jrk[$q], $format14);  # 写入 marker 名

            }
            # 遍历写入 "术前患者" "术前供者" "术后外周血/术后骨髓" "术后T细胞" "术后B细胞" "术后NK细胞" "术后粒细胞" 的型别信息
            for my $q (0..$#markers_jrk){
                # $worksheet->write($q+19,0,$markers_jrk[$q], $format15);  # 写入 marker 名
                # 写入 术前患者 的型别
                if (%hash_this_patient_ID_and_shuqian_patient_genotypes){
                    $worksheet->write($q+19,1,decode('GB2312',$hash_this_patient_ID_and_shuqian_patient_genotypes{$tempid}->{$markers_jrk[$q]}), $format15);
                    print "L3213:$q|markers_jrk[$q]|$hash_this_patient_ID_and_shuqian_patient_genotypes{$tempid}->{$markers_jrk[$q]}\n" ;
                } else {
                    $worksheet->write($q+19,1,decode('GB2312',''), $format15);
                    print "L3216:$q|markers_jrk[$q]|\$hash_this_patient_ID_and_shuqian_patient_genotypes{\$tempid}->{\$markers_jrk[\$q]} not exists.\n" ;
                    print "L3217: Error! Please check!\n" ;
                }

                # 写入 术前供者 的型别
                if (%hash_this_patient_ID_and_shuqian_donor_genotypes){
                    $worksheet->write($q+19,2,decode('GB2312',$hash_this_patient_ID_and_shuqian_donor_genotypes{$tempid}->{$markers_jrk[$q]}), $format15);
                    print "L3222:$q|markers_jrk[$q]|$hash_this_patient_ID_and_shuqian_donor_genotypes{$tempid}->{$markers_jrk[$q]}\n" ;
                } else {
                    $worksheet->write($q+19,2,decode('GB2312',''), $format15);
                    print "L3224:$q|markers_jrk[$q]|\$hash_this_patient_ID_and_shuqian_donor_genotypes{\$tempid}->{\$markers_jrk[\$q]} not exists.\n" ;
                    print "L3217: Error! Please check!\n" ;
                }

                # 写入 术后外周血/骨髓血 的型别
                if (%hash_this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_genotypes){
                    $worksheet->write($q+19,3,decode('GB2312',$hash_this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_genotypes{$tempid}->{$markers_jrk[$q]}), $format15);
                    print "L3222:$q|markers_jrk[$q]|$hash_this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_genotypes{$tempid}->{$markers_jrk[$q]}\n" ;
                } else {
                    $worksheet->write($q+19,3,decode('GB2312',''), $format15);
                    print "L3224:$q|markers_jrk[$q]|\$hash_this_patient_ID_and_shuhou_patient_waizhouxue_or_gusuixue_genotypes{\$tempid}->{\$markers_jrk[\$q]} not exists.\n" ;
                }

                 # 写入 术后T细胞 的型别
                if (%hash_this_patient_ID_and_shuhou_patient_T_cell_genotypes){
                    $worksheet->write($q+19,4,decode('GB2312',$hash_this_patient_ID_and_shuhou_patient_T_cell_genotypes{$tempid}->{$markers_jrk[$q]}), $format15);
                    print "L3240:$q|markers_jrk[$q]|,$hash_this_patient_ID_and_shuhou_patient_T_cell_genotypes{$tempid}->{$markers_jrk[$q]}\n" ;
                } else {
                    $worksheet->write($q+19,4,decode('GB2312',''), $format15);
                    print "L3242:$q|markers_jrk[$q]|\$hash_this_patient_ID_and_shuhou_patient_T_cell_genotypes{\$tempid}->{\$markers_jrk[\$q]} not exists.\n" ;
                }

                 # 写入 术后B细胞 的型别
                if (%hash_this_patient_ID_and_shuhou_patient_B_cell_genotypes){
                    $worksheet->write($q+19,5,decode('GB2312',$hash_this_patient_ID_and_shuhou_patient_B_cell_genotypes{$tempid}->{$markers_jrk[$q]}), $format15);
                    print "L3248:$q|markers_jrk[$q]|$hash_this_patient_ID_and_shuhou_patient_B_cell_genotypes{$tempid}->{$markers_jrk[$q]}\n" ;
                } else {
                    $worksheet->write($q+19,5,decode('GB2312',''), $format15);
                    print "L3250:$q|markers_jrk[$q]|\$hash_this_patient_ID_and_shuhou_patient_B_cell_genotypes{\$tempid}->{\$markers_jrk[\$q]} not exists.\n" ;
                }

                 # 写入 术后NK细胞 的型别
                if (%hash_this_patient_ID_and_shuhou_patient_NK_cell_genotypes){
                    $worksheet->write($q+19,6,decode('GB2312',$hash_this_patient_ID_and_shuhou_patient_NK_cell_genotypes{$tempid}->{$markers_jrk[$q]}), $format15);
                    print "L3256:$q|markers_jrk[$q]|$hash_this_patient_ID_and_shuhou_patient_NK_cell_genotypes{$tempid}->{$markers_jrk[$q]}\n" ;
                } else {
                    $worksheet->write($q+19,6,decode('GB2312',''), $format15);
                    print "L3258:$q|markers_jrk[$q]|\$hash_this_patient_ID_and_shuhou_patient_NK_cell_genotypes{\$tempid}->{\$markers_jrk[\$q]} not exists.\n" ;
                }

                 # 写入 术后粒细胞 的型别
                if (%hash_this_patient_ID_and_shuhou_patient_li_cell_genotypes){
                    $worksheet->write($q+19,7,decode('GB2312',$hash_this_patient_ID_and_shuhou_patient_li_cell_genotypes{$tempid}->{$markers_jrk[$q]}), $format15);
                    print "L3264:$q|markers_jrk[$q]|$hash_this_patient_ID_and_shuhou_patient_li_cell_genotypes{$tempid}->{$markers_jrk[$q]}\n" ;
                } else {
                    $worksheet->write($q+19,7,decode('GB2312',''), $format15);
                    print "L3266:$q|markers_jrk[$q]|\$hash_this_patient_ID_and_shuhou_patient_li_cell_genotypes{\$tempid}->{\$markers_jrk[\$q]} not exists.\n" ;
                }

            }

            # 检测者 + 复核者 + 报告日期 + 报告盖章
            # 在检测日期处插入矩形框，设置透明度
            # 在 "报告" 表单 相应位置，插入 "检测者" "复核者" "盖章" 图片
            if (-e "pic/检测者.png"){
                 $worksheet->insert_image('C45', "pic/检测者.png", 0, 4, 1.5, 1.5);
            }
            if (-e "pic/复核者.png"){
                 $worksheet->insert_image('E45', "pic/复核者.png", 0, 8, 1.0, 1.0);
            }
            if (-e "pic/盖章.png"){
                 $worksheet->insert_image('G45', "pic/盖章.png", 3, -30, 0.998, 0.977);
            }

            my $rectangle_report_date = $workbook->add_shape(type => "rect", text => decode('GB2312',sprintf("%d-%02d-%02d",$year,$mon,$mday)), format => $format21, valign => 'ctr', align => 'l', line => '');
            $worksheet->insert_shape('G45', $rectangle_report_date, 0, 0, 2, 1);
            $worksheet->write('B45', decode('GB2312','检测者：'), $format16);
            $worksheet->write('C45', decode('GB2312',''), $format16);
            $worksheet->write('D45', decode('GB2312','复核者：'), $format16);
            $worksheet->write('E45', decode('GB2312',''), $format16);
            $worksheet->write('F45', decode('GB2312','报告日期：'), $format16);
            # $worksheet->merge_range('G45:H45', decode('GB2312',sprintf("%d-%02d-%02d",$year,$mon,$mday)), $format21);


            ################## 备注 部分 #########################
            $worksheet->merge_range('A46:A52', decode('GB2312','备注：'), $format17);
            $worksheet->merge_range('B46:H46', decode('GB2312','1.本检测采用STR-PCR和毛细管电泳片段分析方法。'), $format18);
            $worksheet->merge_range('B47:H47', decode('GB2312','2.本报告仅对本次检验的样本负责，结果仅供临床医生参考。'), $format19);
            $worksheet->merge_range('B48:H48', decode('GB2312','3.样本保存有一定期限，若对报告结果有疑义，请在自报告日期起7天内提出复检申请，逾期不再受理。'), $format19);
            $worksheet->merge_range('B49:H52', decode('GB2312','4.嵌合状态界定[1]:
完全嵌合状态(CC): DC≥95%; 混合嵌合状态(MC): 5%≤DC<95%； 微嵌合状态: DC<5%。
[1] Outcome of patients with hemoglobinopathies given either cord blood or bone marrow transplantation from an HLA-idebtucak sibling.Blood.2013,122(6):1072-1078.'), $format20);

            ########################################### 输出 "报告" 表单 完成 ###################################################################
            ##################################################################################################################################

            ##################################################################################################################################
            ########################################### 输出 "temp" 表单 (用于绘制嵌合曲线) ###################################################################
            # $tempid = $identity{$TCAID};  # 获取 报告单编号 对应的 患者编码 # %identity 存储 报告单编号 <=> 患者编码(HUN001胡琳)
            my $i;
            my $j = 1;
            my $Chart_Marker_Num = 0;
            my %Graphic_SampleID;
            my %Graphic_Chimerism;
            my %Types;
            my @date_seq;
            push @date_seq, 0;
            if(exists $Chimerism{$tempid}){  #  判断 患者编码 对应的嵌合结果 在 %Chimerism 里是否存在 # %Chimerism 将该 患者编码 对应的嵌合结果（有效位点嵌合率的平均值）存入$Chimerism{$tempid}
                    foreach $i(0 .. $#{$Chimerism{$tempid}}){  # 遍历 患者编码 对应的几份报告的嵌合率
                            my $Chmrsm = $Chimerism{$tempid}[$i];  # "患者编码" 对应的第几份报告的 嵌合率
                            $Chmrsm =~ s/%//;  # 嵌合率是 百分比格式
                            $Chmrsm = sprintf ("%.2f", $Chmrsm);  # 将 百分比格式 转换为 小数（保留小数点后2位）
                            my $Smplid = $SampleID{$tempid}[$i];   # 实验编码 # 数组里每一个元素都是 hash表，用于存储每个 患者编码 对应的每次检测的 实验编码
                            my $SmpType = $sampleType{$Smplid};  # 供患信息中 实验编码 对应的 分选类型 / 样本类型
                            next unless $SmpType;  # 如果样本类型为空，则跳到下一份报告单
                            next if $SmpType eq "-";  # 如果样本类型为 "-"，则跳到下一份报告单
                            my $rptDate = DateUnify($ReportDate{$tempid}[$i]);  # 获取 患者编码 对应的 第i份 报告的 报告日期
                            my $rcvDate = DateUnify($receiveDate{$Smplid});  # 获取 患者编码 对应的 第i份 报告的 收样日期
                            my $smplDate = DateUnify($sampleDate{$Smplid});  # 获取 患者编码 对应的 第i份 报告的 采样日期
                            my $tmpDate;
                            if ($rcvDate ne '不详' && $rcvDate ne '-'){  # 收样日期 不为 "不详" 或 "-"
                                    $tmpDate = $rcvDate;
                            }elsif ($smplDate ne '不详' && $smplDate ne '-'){  # 采样日期 不为 "不详" 或 "-"
                                    $tmpDate = $smplDate;
                            }elsif($rptDate ne '不详'){  # 报告日期 不为 "不详"
                                    $tmpDate = $rptDate;
                            }else{  # 采样日期 / 收样日期 / 报告日期 都是 "不详" 或者 "-"
                                    $tmpDate = sprintf "%s%d%s", "术后", $j, "次";
                            }

                            # print $Smplid,"|", $rptDate,"|",$rcvDate,"|",$smplDate,"|",$tmpDate,"\n";

                            $Graphic_Chimerism{$tmpDate}{$SmpType} = $Chmrsm;  # 存储 样本 某个时间（采样日期 / 收样日期 / 报告日期）对应的 某个类型的样本 对应的 嵌合率结果
                            $Graphic_SampleID{$tmpDate}{$SmpType} = $Smplid;  # 存储 样本 某个时间（采样日期 / 收养日期 / 报告日期）对应的 某个类型的样本 对应的 实验编码
                            $Types{$SmpType} ++;  # print "L2649:$tempid|$TCAID|" . $SmpType . ":" . $Types{$SmpType} . "\n" ;# 样本类型 / 分选类型 个数递增 (这里 $Types{$SmpType} 没有判断是否存在，也没有设置初始化的值？ 默认初始化的值为 0 ? 测试一下看看)
                            if ($date_seq[-1] ne $tmpDate || $tmpDate =~ /术后/){  # 将样本的 （采样日期 / 收样日期 / 报告日期）存入 @date_seq
                                    push @date_seq, $tmpDate;
                                    $j ++;
                            }
                            $Chart_Marker_Num ++;  # 记录总的报告数？
                    }

                    ############################ 写入 temp 表单 #######################################
                    shift @date_seq;
                    my $headings;
                    push @{$headings}, decode('GB2312', '检测日期');  # 设置为 收样日期
                    my $write_data;
                    foreach (@date_seq){
                            push @{${$write_data}[0]}, decode('GB2312', $_);
                    }
                    $i = 1;
                    foreach (keys %Types){
                            push @{$headings}, decode('GB2312', $_);
                            foreach my $tmp(@date_seq){
                                    if (exists $Graphic_Chimerism{$tmp}{$_}){
                                            push @{${$write_data}[$i]}, decode('GB2312', $Graphic_Chimerism{$tmp}{$_});
                                    }else{
                                            push @{${$write_data}[$i]}, undef;
                                    }
                            }
                            $i ++;
                    }

                    $graphic_temp -> write('A1', $headings);  # 写 表头 列
                    $graphic_temp -> write('A2', $write_data);  # 写入 每次检测的嵌合率结果


                    ############################ 输出 "嵌合曲线" 表单 #######################################
                   #########################  chart的格式 ############################
                    my $Gfmt1 = $workbook->add_format(size => 11, align => 'right', font => decode('GB2312','宋体'));  # chart患者姓名
                    my $Gfmt2 = $workbook->add_format(size => 11, bold => 1, align => 'center', font => decode('GB2312','宋体')); # chart姓名
                    my $Gfmt3 = $workbook->add_format(size => 11, bold => 1, align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1);# chart表头
                    my $Gfmt4 = $workbook->add_format(size => 11, align => 'center', valign => 'vcenter', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1 ); # 检测次数, 检测日期，嵌合率等
                    my $Gfmt5 = $workbook->add_format(size => 11, align => 'left', font => decode('GB2312','宋体') ); # "嵌合体定期检测建议"
                    my $Gfmt6 = $workbook->add_format(size => 10, align => 'left', font => decode('GB2312','宋体'), text_wrap => 1); # "嵌合体定期检测建议"
                    my $Gfmt7 = $workbook->add_format(size => 10, align => 'center', font => decode('GB2312','宋体'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # chart表

                    $Gfmt7->set_num_format('0.00');
                    #####
                    $graphic->hide_gridlines();  # 隐藏 网格线
                    $graphic->keep_leading_zeros();  # 保留 开头的0
                    # 设置各列的宽度
                    $graphic->set_column(0,0,11);
                    $graphic->set_column(1,7,11);
                    # 设置各行的高度
                    my @rows = (45,1,1,28,5,18,8,15.6,15.6,18.6,
                                25.8,18.6,19.2,24.6,19.8,16.2,16.2,16.2,16.2,16.2,
                                16.2,16.2,16.2,24,18.75,16.25,15,15,15);
                    for my $i(0 .. $#rows){
                            $graphic->set_row($i, $rows[$i]);
                    }

                    # 设置 "嵌合曲线" 表单 下面 每次检测结果表格的行高
                    foreach $i(1 .. $Chart_Marker_Num){
                            $graphic->set_row($i+30, 13.5);
                    }

                    # 设置 左右和上 边距
                    $graphic->set_margin_left(0.3);
                    $graphic->set_margin_right(0.3);
                    $graphic->set_margin_top(0.25);
                    $worksheet->set_margin_bottom(0.25);

                    # 设置表单的 页脚
                    # my $footer_graphic = '&L'.decode('GB2312','公司地址：广州市黄埔区新瑞路6号二栋2层A205房、A206房')."\n".
                                 '&L'.decode('GB2312','咨询电话：020-62313880').
                    #             #'&R'.decode('GB2312',$this_patient_ID_and_patient_name{$tempid}.'，第').'&P'.decode('GB2312','页/共').'&N'.decode('GB2312','页');
                                  '&C&9'.decode('GB2312','&P'.'/'.'&N');
                    my $footer_graphic = '&C&9'.decode('GB2312','&P'.'/'.'&N');
                    $graphic->set_footer($footer_graphic);

                    # B1 插入 logo
                    # 在第一行第二列插入 公司logo (pic/logo.png)
                    $graphic->insert_image('A1', "pic/logo.png", 1, 16, 0.17623, 0.176);

                    # 写入 各项信息
                    ################## 报告头 部分 #########################
                    # $graphic->merge_range('A1:H1', decode('GB2312','广州君瑞康生物科技有限公司'), $format1);
                    my $bold = $workbook->add_format(size => 12, bold => 1 , font => decode('GB2312','宋体'));
                    my $normal = $workbook->add_format(size => 11, font => decode('GB2312','宋体'));
                    my $fmt_telephone = $workbook->add_format(size => 10, font => decode('GB2312','宋体'));
                    my $fmt_align = $workbook->add_format( align => 'right', valign => 'top', 'bottom' => 2);
                    $fmt_align->set_text_wrap();
                    $graphic->merge_range_type( 'rich_string', 'A1:H1', $bold, decode('GB2312',"广州君瑞康生物科技有限公司"), "\n", $normal, decode('GB2312',"广州市黄埔区新瑞路6号二栋A205"), "\n", $fmt_telephone, decode('GB2312',"020-62313880"), "\n", $fmt_align);
                    $graphic->merge_range('A2:H2', decode('GB2312',''), $format2);  # 留 空行
                    $graphic->merge_range('A3:H3', decode('GB2312',''), $format2);  # 留 空行
                    $graphic->merge_range('A4:H4', decode('GB2312','移植后嵌合体状态曲线图'), $format3);  #

                    # 写入 "嵌合曲线" 表头
                    # 插入一个空行
                    $graphic->merge_range('A5:H5', decode('GB2312',''), $format2);

                    # 写入 "患者姓名" 表头
                    $graphic->write('A6',decode('GB2312','患者姓名：'), $Gfmt1);
                    # 写入 "患者姓名" 内容 $name[$num3[$z]]
                    $graphic->write('B6',decode('GB2312',$this_patient_ID_and_patient_name{$tempid}), $Gfmt2);
                    # 写入 "样本编号" 表头 及 内容 （实际为 供患信息中的 "实验编码"）
                    # $graphic->write('F6',decode('GB2312','样本编号：'), $Gfmt1);
                    # $graphic->write('G6',decode('GB2312',$number{$num3[$z]}), $Gfmt1);

                    # 插入嵌合曲线图
                    my $chart = $workbook->add_chart(type => 'line', embedded => 1 );
                    my $row_max = $#{${$write_data}[0]}+1; # 获取 "temp" 表单的总行数， +1 是因为 add_series 里是以1为起始的
                    my $col_max = $#{$write_data};  # 获取 "temp" 表单的总列数
                    for my $i(1..$col_max){  # 遍历加入  每个 样本类型
                            my $formula = sprintf "=temp!\$%s1", chr($i+65);  # 格式后续测试时关注
                            $chart->add_series(  # 根据 "temp" 表单，选取画图的数据区域
                                    categories => ['temp', 1,$row_max, 0 , 0],  # 选取 "temp" 的第一列作为 categories
                                    values     => ['temp', 1, $row_max, $i, $i],  # 选择 "temp" 的第i+1列 作为 values (对应每一种样本类型)
                                    name_formula => $formula,
                                    marker   => {  # 设置每个 series (样本类型) 的符号
                                            type    => 'automatic',
                                            size    => 1,
                                    },
                                    data_labels => {
                                        # value => 1,
                                        font  => { name => decode('GB2312','宋体'), size => 10 }
                                    },
                            );
                    }

                    #
                    $chart->set_title( none => 1 );  # 不显示图表的标题
                    $chart->set_chartarea(  # is used to set the properties of the chart area
                            color => 'white',
                            line_color => 'black',
                            line_weight => 2,
                    );

                    $chart->set_plotarea(  # is used to set properties of the plot area of a chart.
                            color => 'white',

                    );

                    $chart->set_y_axis(  # is used to set properties of the Y axis.
                            name => decode('GB2312','嵌合率(%)'),
                            name_font => { name => decode('GB2312','宋体'), size => 10 },
                            num_font => {name => decode('GB2312','宋体'), size => 10},
                            min  => 0,
                            max  => 100,
                            major_unit => 20,
                    );

                    $chart->set_x_axis(  # is used to set properties of the x axis.
                            # name => decode('GB2312','嵌合率(%)'),
                            name_font => { name => decode('GB2312','宋体'), size => 10 },
                            num_font => {name => decode('GB2312','宋体'), size => 10},
                    );

                    # 设置 图例的位置(bottom) / 整个图表的宽高(in pixels)
                    $chart->set_legend( position => 'bottom' , font => {name => decode('GB2312','宋体'), size => 10} );
                    $chart->set_size( width => 650, height => 420 );
                    # 将图表插入 "嵌合曲线" 的 "A8"
                    $graphic->insert_chart('A8', $chart);

                    # 写入 "嵌合曲线" 最下部分 "每次检测的 采样日期/检测日期/嵌合率(%)/样本编号(实际为实验编码)/样本类型" 汇总表
                    # 写入表头
                    $graphic->merge_range('A27:A28', decode('GB2312','检测次数'), $Gfmt3);
                    $graphic->merge_range('B27:B28', decode('GB2312','检测时间'), $Gfmt3);
                    $graphic->merge_range('C27:H27', decode('GB2312','嵌合率(%)'),   $Gfmt3);
                    $graphic->write('C28', decode('GB2312','骨髓血'), $Gfmt3);
                    $graphic->write('D28', decode('GB2312','外周血'), $Gfmt3);
                    $graphic->write('E28', decode('GB2312','T细胞'), $Gfmt3);
                    $graphic->write('F28', decode('GB2312','B细胞'), $Gfmt3);
                    $graphic->write('G28', decode('GB2312','NK细胞'), $Gfmt3);
                    $graphic->write('H28', decode('GB2312','粒细胞'), $Gfmt3);

                    my $i = 0;
                    # my $j = 0;
                    for my $tmpDate(@date_seq){  # 遍历每个采样日期
                            $graphic->write($i+28, 0, $i+1, $Gfmt4);  # 写入 检测次数
                            # 先写一遍空行
                            for(my $j=0;$j<6;$j++){
                                $graphic->write($i+28, 2+$j, '', $Gfmt4);
                            }

                            my $if_smplDate_written_into_sheet = 0 ;
                            for my $SmpType(keys %Types){  # 遍历每个采样日期的每种 样本类型
                                    my $Smplid = $Graphic_SampleID{$tmpDate}{$SmpType};  # 某个采样日期的某个样本类型 对应的 "实验编码"
                                    my $Chmrsm = $Graphic_Chimerism{$tmpDate}{$SmpType};  # 某个采样日期的某个样本类型 对应的 "嵌合率"
                                    next unless $Smplid;  # 跳过 "实验编码" 为空的行
                                    my $rcvDate = $receiveDate{$Smplid};  # 获取 "实验编码" 对应的 "收样日期"
                                    # my $smplDate = $sampleDate{$Smplid};  # 获取 "实验编码" 对应的 "采样日期"
                                    if ($if_smplDate_written_into_sheet == 0){
                                        $graphic->write($i+28, 1, decode('GB2312',$rcvDate), $Gfmt4);  # 写入 "收样日期" （对应表格中的 "检测日期"）
                                        $if_smplDate_written_into_sheet ++ ;
                                    }

                                    # 获取 骨髓血样本的 嵌合率
                                    if ($SmpType =~ /骨髓血|骨髓/){
                                        $graphic->write($i+28, 2, sprintf("%.2f",$Chmrsm), $Gfmt4);  # 写入 "嵌合率" 结果，按百分比个数，保留小数点后2位有效数字
                                    } elsif ($SmpType =~ /外周血/){
                                        $graphic->write($i+28, 3, sprintf("%.2f",$Chmrsm), $Gfmt4);  # 写入 "嵌合率" 结果，按百分比个数，保留小数点后2位有效数字
                                    } elsif ($SmpType =~ /T细胞/){
                                        $graphic->write($i+28, 4, sprintf("%.2f",$Chmrsm), $Gfmt4);  # 写入 "嵌合率" 结果，按百分比个数，保留小数点后2位有效数字
                                    }  elsif ($SmpType =~ /B细胞/){
                                        $graphic->write($i+28, 5, sprintf("%.2f",$Chmrsm), $Gfmt4);  # 写入 "嵌合率" 结果，按百分比个数，保留小数点后2位有效数字
                                    }  elsif ($SmpType =~ /NK细胞/){
                                        $graphic->write($i+28, 6, sprintf("%.2f",$Chmrsm), $Gfmt4);  # 写入 "嵌合率" 结果，按百分比个数，保留小数点后2位有效数字
                                    } elsif ($SmpType =~ /粒细胞/){
                                        $graphic->write($i+28, 7, sprintf("%.2f",$Chmrsm), $Gfmt4);  # 写入 "嵌合率" 结果，按百分比个数，保留小数点后2位有效数字
                                    } else {
                                        print "L3505: Error sample type=$SmpType.Please Check!\n";
                                    }

                                    # $j ++;
                                    # if (($j-11)%54 == 0){  # 当检测次数超过5次时，重新写一次表头
                                    #        $graphic->write($j+31, 1, decode('GB2312','检测次数'), $Gfmt3);
                                    #        $graphic->write($j+31, 2, decode('GB2312','采样日期'), $Gfmt3);
                                    #        $graphic->write($j+31, 3, decode('GB2312','检测日期'), $Gfmt3);
                                    #        $graphic->write($j+31, 4, decode('GB2312','嵌合率(%)'),   $Gfmt3);
                                    #        $graphic->write($j+31, 5, decode('GB2312','样本编号'), $Gfmt3);
                                    #        $graphic->write($j+31, 6, decode('GB2312','样本类型'), $Gfmt3);
                                    #        $j ++;
                                    #}
                            }
                            $i ++;
                    }

                    # "嵌合曲线" 图表下面 提示信息 部分的内容
                    # 写入 "TCA定期检测流程 .... 等说明性 内容"
                    $graphic->merge_range(@date_seq+30, 0, @date_seq+30, 7 ,decode('GB2312','嵌合体定期检测建议：'), $Gfmt5);
                    $graphic->merge_range(@date_seq+31, 0, @date_seq+32, 7, decode('GB2312','1、术后2周进行首次嵌合体检测，第4周进行第二次检测；术后6个月内，每月进行一次检测，6个月后，每2个月检测一次，直至嵌合率稳定。'), $Gfmt6);
                    $graphic->merge_range(@date_seq+33, 0, @date_seq+33, 7, decode('GB2312','2、若术后免疫治疗方案调整，在调整后2周重新进行检测，频率参照以上建议。'),$Gfmt6);
                    # $graphic->merge_range('B30:G30', decode('GB2312','温馨提示：一旦术后免疫治疗方案调整，在调整后2周需要重新启动检测'), $Gfmt6);
            }else{

            }

            # excel 文件写完，关闭文件
            $workbook->close();
            $RptBox -> Append("汇总报告 生成成功！\r\n");  # 更新 "生成报告" 部分的文本框中显示的提示信息

            $tt ++ ;
        }

        # 输出 状态 提示信息 ($sb))
        $sb->Move( 0, ($main->ScaleHeight() - $sb->Height()) );
        $sb->Resize( $main->ScaleWidth(), $sb->Height() );
        if ($success){
                $sb->Text("汇总报告输出完成");
        }else{
                $sb->Text("汇总报告输出完成（有错误）");
        }

        $RUNwindow -> Hide();
        if ($success){
                $error =  "汇总报告输出保存成功！\n";
                Win32::MsgBox $error, 0, "成功！";
        }else{
                $error =  "汇总报告输出保存成功，但发生了错误！\n";
                Win32::MsgBox $error, 0, "注意！";
        }
}

###################################################################################
# RUN_MouseMove 函数: "生成报告" 按钮 鼠标移上 时的处理函数                            #
###################################################################################
sub RUN_MouseMove{
        $sb -> Text('运行，产生分析报告');
}

###################################################################################
# RUN_MouseOut 函数: "生成报告" 按钮 鼠标移出 时的处理函数                            #
###################################################################################
sub RUN_MouseOut{
        $sb -> Text('');
}

###################################################################################
# QUIT_MouseMove 函数: "退出" 按钮 鼠标移上 时的处理函数                              #
###################################################################################
sub QUIT_MouseMove{
        $sb -> Text('退出');
}

###################################################################################
# QUIT_MouseMove 函数: "退出" 按钮 鼠标移出 时的处理函数                              #
###################################################################################
sub QUIT_MouseOut{
        $sb -> Text('');
}

###################################################################################
# QUIT_MouseMove 函数: "退出" 按钮 点击 时的处理函数                                 #
###################################################################################
sub QUIT_Click{
        &WriteConfig;

        return -1;
}

###################################################################################
# COPY_MouseMove 函数: "复制" 按钮 鼠标移上 时的处理函数                              #
###################################################################################
sub COPY_MouseMove{
        $sb -> Text('复制右侧框中所有记录至剪贴板');
}

###################################################################################
# COPY_MouseOut 函数: "复制" 按钮 鼠标移出 时的处理函数                               #
###################################################################################
sub COPY_MouseOut{
        $sb -> Text('');
}

###################################################################################
# COPY_MouseOut 函数: "复制" 按钮 点击 时的处理函数                                   #
###################################################################################
sub COPY_Click{
        $RptBox -> SelectAll();
        $RptBox -> Copy();
        $error = '成功复制至剪贴板！';
        Win32::MsgBox $error, 0 ,"成功！";
}

###################################################################################
# Shorten 函数: 截断显示文件名的函数，在 供患信息列表框 和 已有数据列表框中调用           #
###################################################################################
sub Shorten{
        my ($string, $lim) = @_; print "L2700:" . $string . "\t" . $lim ."\n";
        if ($string =~ /$pwd\\(.+)$/){
                $string = ".\\".$1; print "L2702:" . $string . "\n";
        }
        my $len = length($string);
        return $string if $len <= $lim;
        $string =~ /^.+\\([^\\]+)$/;
        $len = length($1);
        my $tmp = sprintf "%s...\\%s", substr($string, 0 ,$lim-$len-4), $1;
        return $tmp;
}

###################################################################################
# DelItem 函数: 从数组中删除指定下标的元素，在 XXX 中调用                              #
###################################################################################
sub DelItem{
        my $index = pop;
        my @array = @_;
        my $i;
        my @tmp;
        foreach $i(0..$index-1){
                push @tmp, $array[$i];
        }
        foreach $i($index+1 .. $#array){
                push @tmp, $array[$i];
        }

        # @array = ();
        # push @array, $_ for @tmp;
        # return @array;
        return @tmp;
}

###################################################################################
# ReadConfig 函数: 读取配置文件                                                     #
###################################################################################
sub ReadConfig{
        unless (open IN,"TCAconfig.ini"){
                open IN,"> TCAconfig.ini";
                foreach (@ConfigList){
                        print IN $_,"\t",$pwd,"\n";
                }
                close IN;
                `attrib +r +h TCAconfig.ini`;
        }
        while(<IN>){
                next if /^#/;
                chomp;
                my @str = split;
                next unless exists $ConfigHash{$str[0]};
                $ConfigHash{$str[0]} = $str[1];
        }
        close IN;
}

###################################################################################
# WriteConfig 函数: 写入配置文件                                                    #
###################################################################################
sub WriteConfig{
        `attrib -r -h TCAconfig.ini` if (-e "TCAconfig.ini");
        open IN,"> TCAconfig.ini";
        foreach (@ConfigList){
                print IN $_,"\t",$ConfigHash{$_},"\n";
        }
        close IN;
        `attrib +r +h TCAconfig.ini`;
}

###################################################################################
# DateUnify 函数: 日期格式化                                                        #
###################################################################################
sub DateUnify{
        return $_[0] if $_[0] eq '不详';

        if ($_[0] =~ /^\d{8}$/){
                return substr($_[0], 0, 4)."-".substr($_[0], 4, 2)."-".substr($_[0], 6, 2);
        }elsif($_[0] =~ /^(\d+)\/(\d+)\/(\d+)$/){
                return $1."-".sprintf("%02d",$2)."-".sprintf("%02d",$3);
        }elsif($_[0] =~ /^(\d+)-(\d+)-(\d+)$/){
                return $1."-".sprintf("%02d",$2)."-".sprintf("%02d",$3);;
        }else{
                return $_[0];
        }
}

###################################################################################
# Avg_SD 函数: 输入数组，计算数组的均值 和 SD                                         #
###################################################################################
sub Avg_SD{
        my $total = 0;
        my $SD = 0;

        $total += $_ foreach @_;
        $total /= @_;

        $SD += ($total-$_)*($total-$_) foreach @_;
        $SD = sqrt($SD/@_);

        return ($total, $SD);
}