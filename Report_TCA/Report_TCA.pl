#!/usr/bin/perl -w
#20171001 ���»������д�����ǰ����Բ��ϣ�������������ı���
#20171105 ���ڸ�ʽunifying��Ԥ����Ϊ������ţ�Ƕ����ͳһ2λС��
#20170202 ��SD��CV��ҽԺ���ƿ����Զ�����

use strict;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtUnicode;
# use Unicode::Map;
# use Spreadsheet::WriteExcel;
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Chart::Line;
use Encode;
use Win32;
use Win32::GUI();
# use Win32::GUI::Grid;

my ($mday,$mon,$year) = (localtime)[3..5];
$mday = sprintf("%d", $mday);
$mon  = sprintf("%d", $mon + 1);
$year = $year + 1900;

# my $dateXXX = sprintf ("%4d%02d%02d", $year,$mon,$mday);
# my $TrialLim = sprintf ("%d%d%d%d%d%d%d%d", ord('C')-65,ord('A')-65,ord('B')-65,ord('I')-65,ord('A')-65,ord('E')-65,ord('D')-65,ord('A')-65);
my $version = "TCA�����Զ�������� v1.1 (��ʽ��)";

my $pwd = `cd`;
chomp $pwd;

my %sampleDate;
my %receiveDate;
my %sampleType;
my %Chimerism;
my %SampleID;
my %HasChimerism;
my %ReportDate;
my @exp_list;
my $i = 0;
my @data_in;
my %exp_id;
my @TCA_id;
my %exp_num;
my %exp_seq;
my %exp_error;
my %together;
my %identity;
my %history;
my $test = 0;
my $error;
my $InputLoaded = 0;
my $SummaryLoaded = 0;
my $ExpLoaded = 0;

my %HospitalName;
my %HospitalAlias;
my %region;
my %ID;
my %alias;

my $DOS = Win32::GUI::GetPerlWindow();
Win32::GUI::Hide($DOS);

# if ($dateXXX > $TrialLim){
        # $error = '���԰汾����������'.$TrialLim.'
# ����ϵyewei@catb.org.cn';
        # Win32::MsgBox ($error, 0, "�ѹ���");
        # exit(0);
# }

# $error = '�˰汾���ܴ��ڴ��󣬽�������ʹ��
# �κ�����ͽ�������ϵyewei@catb.org.cn
# �Ƿ������';

# my $goon = Win32::MsgBox ($error, 4, "����");

# exit(0) if $goon == 7;

my %allele;
my %PrevAllele;
my $curr_index;

my %ThisAllele;
my %area;
my %trans; #��������ʵ�������д��ȫ�Ƶ�ת��

my @markers = ('D8S1179','D21S11','D7S820','CSF1PO','D3S1358','D5S818','D13S317','D16S539','D2S1338','D19S433','VWA','D12S391','D18S51','Amel','D6S1043','FGA');
my %markerExist;
foreach (@markers){
        $markerExist{$_} = 'yes',
}

my @headers = ('��������', '����ʱ��', '��ֲ����', '��Ʒ����', '��Ʒ����', '��ѡ���', '��������', 'ʵ�����', '���浥���', '����', '������ϵ', '�Ա�', '����', '���', '��Ե��ϵ', '�����������', 'ҽԺ����', '�ͼ�ҽԺ', '�ͼ�ҽ��', 'סԺ��', '����');

my $InputSample_str = "δѡ��";
my $ThisExp_str = "δѡ��";
my $PrevExp_str = "δѡ��";
my $Output_Dir = "δѡ��";
my $Output_rpt_str = "δѡ��";
my $Output_calc_str = "δѡ��";

my $output_name;

my @InputFound;
my @PrevFound;
my @InputList;
my @PrevList;
my @ThisFound;
my @ThisList;

my @ConfigList = ("InputLoc", "SummaryLoc", "ThisLoc", "OutputLoc");
my %ConfigHash;

foreach (@ConfigList){
                $ConfigHash{$_} = $pwd;
}

&ReadConfig;

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

my $sb = $main->AddStatusBar();

my $text11 = $main->AddLabel(
        -text => '������Ϣ',
        -pos => [10, 10],
        -font => $font,
);

my $text12 = $main->AddLabel(
        -text => '���ҵ�',
        -pos => [10, 45],
);

my $Input1 = $main-> AddListbox(
        -name => "List1",
        -left => 10,
        -top => 60,
        -width => 250,
        -height => 80,
        -hscroll => 1,
        -vscroll => 1,
);

my $open1 = $main->AddButton(
        -name => "Open1",
        -text => "����λ��...",
        -pos  => [ 10, 130 ],
);

my $text14 = $main->AddLabel(
        -pos => [140, 135],
        -text => '��˫����ȡ',
);

my $text13 = $main->AddLabel(
        -pos => [10, 155],
        -width => 250,
        -text => '��δ��ȡ',
);

my $text21 = $main->AddLabel(
        -text => '��������',
        -pos => [300, 10],
        -font => $font,
);

my $text22 = $main->AddLabel(
        -text => '���ҵ�',
        -pos => [300, 45],
);

my $Input2 = $main-> AddListbox(
        -name => "List2",
        -left => 300,
        -top => 60,
        -width => 250,
        -height => 80,
        -hscroll => 1,
        -vscroll => 1,
);

my $open2 = $main->AddButton(
        -name => "Open2",
        -text => "����λ��...",
        -pos  => [ 300, 130 ],
);

my $text24 = $main->AddLabel(
        -pos => [430, 135],
        -text => '��˫����ȡ',
);

my $text23 = $main->AddLabel(
        -pos => [300, 155],
        -width => 250,
        -text => '��δ��ȡ',
);

my $display2 = $main->AddButton(
        -name => "DISPLAY2",
        -text => "��ӡ���з���",
        -pos  => [ 10, 180 ],
        -size => [ 545 , 30],
        -disabled => 1,
);

# my $dpwindow = new Win32::GUI::Window (
        # -name  => "W2",
        # -title => "�鿴���з���",
        # -pos   => [ 300, 300 ],
        # -size  => [ 400, 700 ],
        # -parent => $main,
        # -sizabke => 0,
        # -resizable => 0,
# );
# my $Grid = new Win32::GUI::Grid (
    # -parent  => $dpwindow,
    # -name    => "Grid",
    # -pos     => [0, 0],
# ) or die "new Grid";

# $Grid->SetEditable(0);
# $Grid->SetRows(18);
# $Grid->SetColumns(3);
# $Grid->SetFixedRows(1);
# $Grid->SetFixedColumns(1);

# my $w2_prev = $dpwindow->AddButton(
        # -name => "W2_PREV",
        # -text => "��һ��",
        # -pos  => [ 50, 600 ],
# );

# my $w2_next = $dpwindow->AddButton(
        # -name => "W2_NEXT",
        # -text => "��һ��",
        # -pos  => [ 150, 600 ],
# );

# my $w2_close = $dpwindow->AddButton(
        # -name => "W2_CLOSE",
        # -text => "�ر�",
        # -pos  => [ 250, 600 ],
# );

my $open3 = $main->AddButton(
        -name => "Open3",
        -text => "����»�����",
        -pos  => [ 10, 220 ],
        -size => [ 100, 30 ],
        -disabled => 1,
);

my $del3 = $main->AddButton(
        -name => "Del3",
        -text => "�Ƴ�",
        -pos  => [ 80, 250 ],
        -size => [ 30, 20 ],
        -disabled => 1,
);

my $Read3 = $main->AddButton(
        -name => "Read3",
        -text => "��ȡ",
        -pos  => [ 500, 221 ],
        -size => [ 50, 50 ],
        -disabled => 1,
);

my $Input3 = $main-> AddListbox(
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

my $text3 = $main->AddLabel(
        -text => "��δ��ȡ",
        -pos => [120, 280],
);

my $sep = $main-> AddLabel(
        -text =>"==============================================================================================================================",
        -pos => [0,300],
);

my $run4 = $main->AddButton(
        -name => "RUN",
        -text => "���ɱ���",
        -font => $font,
        -pos  => [ 20, 320 ],
        -size => [160,60],
        -disabled => 1,
);

my $RptBox = $main-> AddTextfield(
        -name => "RptBox",
        -pos => [200, 310],
        -size => [350, 135],
        -readonly => 1,
        -multiline => 1,
        -vscroll => 1,
        -autovscroll => 1,
        -autohscroll => 0,
);

# my $dpsb2 = $dpwindow->AddStatusBar();

my $RUNwindow = new Win32::GUI::Window (
        -name  => "RUNWindow",
        -title => "���������ļ������Ժ�...",
        -pos   => [ 300, 300 ],
        -size  => [ 300, 100 ],
        -parent => $main,
        -sizabke => 0,
        -resizable => 0,
);

my $quit = $main->AddButton(
        -name => "QUIT",
        -text => "�˳�",
        -pos  => [ 20, 420 ],
        -size => [ 60,20],
);

my $copybutton = $main -> AddButton(
        -name => "COPY",
        -text => "����",
        -pos  => [120, 420],
        -size => [ 60,20],
);

# my $checkbox => $main -> AddCheckbox(
        # -name => "UseHuatuo",
        # -pos  => [20, 390],
# );

my $direct = 1;
$main->Show();

$sb -> Text('���ڶ�ȡҽԺ��Ϣ');
unless (open IN,".HospitalTrans.txt"){
        $error = "δ�ҵ�ҽԺ��Ϣ
��������½�ѹ��װ����";
        my $s = Win32::MsgBox $error,1, "����";
        exit(0);
}
<IN>;
while(<IN>){
        my @str = split;
        $region{$str[2]} = $str[0];
        $ID{$str[2]} = $str[1];
        while ($#str > 2){
                my $tmp = pop @str;
                $alias{$tmp} = $str[2];
        }
}
close IN;

$sb -> Text('���ڼ������ļ�');
unless (-e ".PrevSamples.txt"){
        $error = "δ�ҵ�����������Ϣ
��������½�ѹ��װ����";
        my $s = Win32::MsgBox $error,1, "����";
        exit(0);
}
open IN,".PrevSamples.txt";
<IN>;

while(<IN>){
        chomp;
        my @str = split /\t/, $_;
# 0���� 1��ݵ��� 2 3 4�������� 5�������� 6��ֲ���� 7��Ʒ���� 8��Ʒ���� 9��Ʒ���� 10������� 11���浥��� 12���� 13������ϵ 14�Ա� 15���� 16��� 17��Ե��ϵ 18����������� 19ҽԺ���� 20�ͼ�ҽԺ 21�ͼ�ҽ�� 22�ʼı����ַ 23�ʼı����ַ
        next unless $str[10];
        next unless $str[11];
        next unless $str[12];
        next unless $str[19];

        my $Smplid = $str[10];
        $sampleDate{$Smplid} = $str[4]? $str[4]:'����';
        $sampleDate{$Smplid} = DateUnify($sampleDate{$Smplid});
        $receiveDate{$Smplid} = $str[5]? $str[5]:'����';
        $receiveDate{$Smplid} = DateUnify($receiveDate{$Smplid});
        $sampleType{$Smplid} = $str[8];
}

# print "STR1610506 Sample: ", $sampleDate{STR1610506},"\n";
# print "STR1610506 Recieve:", $receiveDate{STR1610506},"\n";

close IN;

unless (-e ".PrevChimerism.txt"){
        $error = "δ�ҵ�����Ƕ������Ϣ
��������½�ѹ��װ����";
        my $s = Win32::MsgBox $error,1, "����";
        exit(0);
}

open IN,".PrevChimerism.txt";
<IN>;
while (<IN>){
        chomp;
#0������        1��������        2�������        3��ع���/����        4Ƕ����        5��������        6ҽԺ���        7ҽԺȫ��        8��ע        9��������        10��Ʒ����
        my @str = split /\t/, $_;
        next unless $str[4] =~ /\d+(\.\d+)?%/;
        next unless $str[2];
        next unless $str[1];
        next if $str[7] =~ /N\/A/;
        next unless $str[7];
        next if $str[7] eq '����';
        next unless $str[6];

        if (exists $ID{$str[7]}){
                $str[6] = $ID{$str[7]};
        }elsif(exists $alias{$str[7]}){
                my $tmp = $alias{$str[7]};
                $str[6] = $ID{$tmp};
                $str[7] = $tmp;
        }

        my $tmp = $str[6].$str[1];
        push @{$Chimerism{$tmp}}, $str[4];
        push @{$SampleID{$tmp}}, $str[2];
        $HasChimerism{$str[2]} = $tmp;
        if ($str[5]){
                push @{$ReportDate{$tmp}}, DateUnify($str[5]);
        }else{
                push @{$ReportDate{$tmp}}, "����";
        }
}
close IN;
# my $teststr = "FJ001�����";
# $i = 0;
# foreach (@{$Chimerism{$teststr}}){
        # print "Ƕ���ʣ�  ",$_,"\n";
        # print "������ţ�",${$SampleID{$teststr}}[$i],"\n";
        # print "����ʱ�䣺",${$ReportDate{$teststr}}[$i],"\n";
        # $i++;
# }

# foreach (keys %Chimerism){
        # print $_,"|", $#{$Chimerism{$_}}+1,"\n";
        # my $i = 0;
        # foreach my $tmp(@{$Chimerism{$_}}){
                # my $rptDate;
                # if ($ReportDate{$_}[$i] eq "����"){
                        # $rptDate = sprintf "%s%d%s", "����", $i+1 , "��";
                # }else{
                        # $rptDate = $ReportDate{$_}[$i];
                # }
                # printf "%15s%15s%20s\n", $rptDate, $SampleID{$_}[$i], $tmp;
                # $i ++;
        # }
# }

#Looking for InputSample files
my $temp = $ConfigHash{InputLoc};
my @filelist = `dir /b $temp\\*.txt`;
my $localnumber = 1;

foreach (@filelist){
        chomp;
        my $tmpfilename = $temp."\\".$_;
        $sb -> Text("���ڶ�ȡ�����ļ�".substr('.....', 0,($localnumber++)%5+1));
        next if /^\./;
        (open IN ,$tmpfilename) || next;
        my $str = <IN>;
        next unless $str;# in case of empty file.
        chomp $str;
        close IN;
#0����ʱ��        1��ֲ����        2��Ʒ����        3��Ʒ����        4��Ʒ����        5��ѡ���        6��������        7ʵ�����        8���浥���        9����        10������ϵ        11�Ա�        12����        13���        14��Ե��ϵ        15�����������        16ҽԺ����        17�ͼ�ҽԺ        18�ͼ�ҽ��
        my $yes = 1;
        my @tmp = split /\t/, $str;
        next if @tmp != 19;
        foreach my $i(0..18){
                $yes = 0 if $tmp[$i] ne $headers[$i];
        }
        next if $yes != 1;
        push @InputFound, $tmpfilename;
        push @InputList, &Shorten($tmpfilename, 39);
}
####

$Input1 -> Add(@InputList);

#Looking for Previous Results files
$temp = $ConfigHash{SummaryLoc};
@filelist = `dir /b $temp\\*.txt`;
$localnumber = 1;
foreach (@filelist){
        chomp;
        my $tmpfilename = $temp."\\".$_;
        $sb -> Text("���ڶ�ȡ�����ļ�".substr('.....', 0,($localnumber++)%5+1));
        next if /^\./;
        (open IN ,$tmpfilename) or next;
        my $str = <IN>;
        next unless $str;# in case of empty file.
        chomp $str;
        close IN;

        my $yes = 1;
        my @tmp = split /\t/, $str;
        next if @tmp != 17;
        foreach my $i(1..16){
                $yes = 0 unless exists $markerExist{$tmp[$i]};
        }
        next if $yes != 1;
        push @PrevFound, $tmpfilename;
        push @PrevList, &Shorten($tmpfilename, 39);
}
####
$Input2 -> Add(@PrevList);
$sb -> Text('');

Win32::GUI::Dialog();
Win32::GUI::Show($DOS);

exit(0);

sub Main_Terminate {

        &WriteConfig;
        return -1;
}

sub List1_DblClick{
        if (@exp_list){

                $error = "�Ѿ��ɹ���ȡ���ݣ��Ƿ����¶�ȡ��";
                my $s = Win32::MsgBox $error,1, "ע�⣡";
                # $Sure = 0;
                # $Msg1 -> DoModal();
                return 0 if $s != 1;
        }

        $InputLoaded = 0;
        $display2->Enable($InputLoaded*$SummaryLoaded);
        $open3->Enable($InputLoaded*$SummaryLoaded);

        $error = "��δ��ȡ";
        $text13 -> Text($error);

        my $sel = $Input1->GetCurSel();
        $InputSample_str = $InputFound[$sel];

        unless (open IN,$InputSample_str){
                $error = "�ļ���ʧ�ܣ�\n";
                Win32::MsgBox $error, 0, "����";
                return 0;
        }
        #0����ʱ��   1��ֲ����   2��Ʒ����   3��Ʒ����   4��Ʒ����   5��ѡ���   6��������   7ʵ�����   8���浥���   9����   10������ϵ   11�Ա�   12����   13���   14��Ե��ϵ   15�����������   16ҽԺ����   17�ͼ�ҽԺ   18�ͼ�ҽ��  19סԺ��   20����

        @exp_list = ();
        $i = 0;
        @data_in = ();
        %exp_id = ();
        @TCA_id = ();
        %exp_num = ();
        %exp_seq = ();
        %exp_error = ();
        %together = ();
        %history = ();
        $test = 0;

        my $tmp = <IN>;
        chomp $tmp;
        #if error content

        my $yes = 1;
        my @tmp = split /\t/, $tmp;
        $yes = 0 if @tmp != 21;
        foreach $i(0..20){
                $yes = 0 if $tmp[$i] ne $headers[$i];
        }

        if ($yes != 1){
                $error = "����ļ�ò�Ʋ��ԡ���ͷӦΪ��\n����ʱ�� ��ֲ���� ��Ʒ���� ��Ʒ���� ��Ʒ���� ��ѡ��� �������� ʵ�����\n���浥��� ���� ������ϵ �Ա� ���� ��� ��Ե��ϵ ����������� ҽԺ����\n�ͼ�ҽԺ �ͼ�ҽ�� סԺ�� ����\n";
                Win32::MsgBox $error, 0, "����";
                return;
        }

        while (<IN>){
                chomp;
                my @str = split /\t/, $_;
                if ($str[7] eq ""){
                        $error = "FATAL!! ʵ�����Ϊ�գ�\n";
                        Win32::MsgBox $error, 0, "����";
                        exit(0);
                }
                $str[7] =~ s/\s+//g;
                if ($str[4] eq "" || $str[8] eq "" || $str[10] eq ""){
                        $error =  "�벹ȫ������Ϣ����Ʒ����/���浥���/������ϵ����".$str[7]."\n";
                        Win32::MsgBox $error, 0, "����";
                        exit(0);
                }
                $str[4] =~ s/\s+//g;
                $str[8] =~ s/\s+//g;
                $str[10] =~ s/\s+//g;
                push @exp_list, $str[7]; #exp_list��InputSamples��˳����ÿһ��ʵ�����
                if (@TCA_id == 0 || $str[8] ne $TCA_id[-1]){
                        push @TCA_id, $str[8];
                }

                $exp_id{$str[7]} = $i; #exp_id ��ϣ�����ÿ��ʵ������ԭʼ˳��
                push @{$together{$str[8]}}, $str[7]; #together ��ά��ϣ����һά�ǹ�ϣ�����Ǳ�����룬�ڶ�ά���б�����ÿ����������Ӧ��ʵ�����
                $exp_num{$str[8]} = @{$together{$str[8]}}; #exp_num һά��ϣ������ÿ��TCA�����а��������������

                ###�����ʵ�����Ŀ####
                if ($str[4] eq "����" && $str[10] eq "����"){

                        if (exists $HasChimerism{$str[7]}){
                                $error = "�������".$str[7]."(".$str[9].")�Ѿ�����Ƕ���ʣ�
���ν���������Ƕ������Ϣ";
##ע���Ǹ��ǣ����Ժ���Ҫɾ��������Ϣ
                                Win32::MsgBox $error, 0, "ע��";
                                my $tmp = $HasChimerism{$str[7]};
                                my $tmpNum;
                                foreach (0 .. $#{$Chimerism{$tmp}}){
                                        $tmpNum = $_ if ${$SampleID{$tmp}}[$_] eq $str[7];
                                }
                                @{$Chimerism{$tmp}} = &DelItem(@{$Chimerism{$tmp}}, $tmpNum);
                                @{$SampleID{$tmp}} = &DelItem(@{$SampleID{$tmp}}, $tmpNum);
                                @{$ReportDate{$tmp}} = &DelItem(@{$ReportDate{$tmp}}, $tmpNum);
                        }
                        $test ++;
                        unless ($str[17]){
                                $history{$str[8]} = 0;
                                $identity{$str[8]} = "NotFound";
                        }else{
                                my $hospital = $str[17];
                                if (exists $ID{$hospital}){
                                        $identity{$str[8]} = $ID{$hospital}.$str[9];
                                }elsif(exists $alias{$hospital}){
                                        my $tmp = $alias{$hospital};
                                        $str[17] = $tmp;
                                        $identity{$str[8]} = $ID{$tmp}.$str[9];
                                }else{
                                        print $str[17],"ľ���ҵ�\n";
                                        $error = $str[17]."ľ���ҵ�
�������ӵ�.HospitalTrans.txt";
                                        Win32::MsgBox $error, 0, "����";
                                        exit(0);
                                        $history{$str[8]} = 0;
                                        $identity{$str[8]} = "NotFound";
                                }
                        }

                        if ($identity{$str[8]} ne "NotFound"){
                                if (exists $Chimerism{$identity{$str[8]}}){
                                        $history{$str[8]} = $#{$Chimerism{$identity{$str[8]}}}+1;
                                }else{
                                        $history{$str[8]} = 0;
                                }
                        }
                }else{
                        if (exists $HasChimerism{$str[7]}){
                                $error = "�������".$str[7]."(".$str[9].")��ʾ����Ƕ���ʣ�
���Ⲣ����һ�������ߵ�������
������ٴ�����";
                                Win32::MsgBox $error, 0, "����";
                                return -1;
                        }
                }
                ##############

                foreach my $tmp(0..20){
                        if ($str[$tmp]){
                                push @{$data_in[$i]}, $str[$tmp]; #data_in ��ά�б���InputSamples��˳�򱣴�ÿһ��ʵ����Ϣ
                        }else{
                                push @{$data_in[$i]}, "-";
                        }
                }
                # print $str[7], "\t", $data_in[$i][14],"\n";
                $i ++;
        }
        close IN;
        print "New Experiments: ",$test,"\n";
        $error = "�Ѷ�ȡ���±�������".$test;
        $text13 -> Text($error);

        foreach my $TCAID(keys %together){
                my @ddd=();
                my @temptogether = ();
        #        print $TCAID,"|",$exp_num{$TCAID},"\n";
                $exp_error{$TCAID} = 0;
                if ($exp_num{$TCAID} == 3){
                        foreach my $exp_str_tmp(@{$together{$TCAID}}){
                                # print "$exp_str_tmp\n";
                                my $i = $exp_id{$exp_str_tmp};
                                if ($data_in[$i][4] eq "��ǰ"){
                                        if ($data_in[$i][10] eq "����"){
                                                $ddd[0] = $i;
                                        }else{
                                                $ddd[1] = $i;
                                        }
                                }else{
                                        if ($data_in[$i][10] eq "����"){
                                                $ddd[2] = $i;
                                        }else{
                                                $error = "�����ţ�".$TCAID."��Ӧ�������������������飡\n";
                                                Win32::MsgBox $error, 0, "����";

        #########################
        # Win32::MsgBox(MESSAGE [, FLAGS [, TITLE]])
        # Create a dialog box containing MESSAGE. FLAGS specifies the required icon and buttons according to the following table:
        #
        # 0 = OK
        # 1 = OK and Cancel
        # 2 = Abort, Retry, and Ignore
        # 3 = Yes, No and Cancel
        # 4 = Yes and No
        # 5 = Retry and Cancel
        #
        # MB_ICONSTOP          "X" in a red circle
        # MB_ICONQUESTION      question mark in a bubble
        # MB_ICONEXCLAMATION   exclamation mark in a yellow triangle
        # MB_ICONINFORMATION   "i" in a bubble
        # TITLE specifies an optional window title. The default is "Perl".
        #
        # The function returns the menu id of the selected push button:
        #
        # 0  Error
        # 1  OK
        # 2  Cancel
        # 3  Abort
        # 4  Retry
        # 5  Ignore
        # 6  Yes
        # 7  No
        ########################

                                                $exp_error{$TCAID} = 1;
                                        }
                                }
                        }
                        next if $exp_error{$TCAID} == 1;
                        $exp_seq{$TCAID} = join ",", @ddd;
                        # print $exp_seq{$TCAID},"\n";
                        # foreach my $i(0..2){
                                # print $ddd[$i],"|",$exp_list[$ddd[$i]],"|",$data_in[$ddd[$i]][4],"|",$data_in[$ddd[$i]][10],"\n";
                        # }
                }elsif ($exp_num{$TCAID} == 2){
                        my @total=();
                        foreach my $exp_str_tmp(@{$together{$TCAID}}){
                                # print "$exp_str_tmp\n";
                                my $i = $exp_id{$exp_str_tmp};
                                my $sum = 0;
                                if ($data_in[$i][4] eq "��ǰ"){
                                        $sum += 0;
                                }else{
                                        $sum += 2;
                                }
                                if ($data_in[$i][10] eq "����"){
                                        $sum += 0;
                                }else{
                                        $sum += 1;
                                }
                                push @total, $sum;
                                push @total, $i;
                        }
                        if ($total[0]> $total[2]){
                                $exp_seq{$TCAID} = join ",", $total[3], $total[1];
                        }else{
                                $exp_seq{$TCAID} = join ",", $total[1], $total[3];
                        }
                        next if $exp_error{$TCAID} == 1;
                        # print $exp_seq{$TCAID},"\n";
                        # my @tmpstr = split ",", $exp_seq{$TCAID};
                        # foreach my $i(@tmpstr){
                                # print $i,"|",$exp_list[$i],"|",$data_in[$i][4],"|",$data_in[$i][10],"\n";
                        # }
                }elsif ($exp_num{$TCAID} == 1){
                        my $exp_str_tmp = ${$together{$TCAID}}[0];
                        # print "$exp_str_tmp\n";
                        my $i = $exp_id{$exp_str_tmp};
                        $error = "�����ţ�".$TCAID."ֻ����һ��ʵ��������Ϊ".$data_in[$i][4].$data_in[$i][10]."������\n";
                        Win32::MsgBox $error, 0, "ע�⣡";
                        $exp_seq{$TCAID} = $i;
                        # print $i,"|",$exp_list[$i],"|",$data_in[$i][4],"|",$data_in[$i][10],"\n";
                }else{
                        $error  = "�����ţ�".$TCAID."��ʵ��������Ź��࣡���飡\n";
                        Win32::MsgBox $error, 0, "ע�⣡";
                        $exp_error{$TCAID} = 1;
                }
        }

        $InputLoaded = 1;
        $display2->Enable($InputLoaded*$SummaryLoaded);
        $open3->Enable($InputLoaded*$SummaryLoaded);
}

sub List1_MouseMove{
        $sb -> Text('˫����Ŀ�Զ�ȡ');
}

sub List1_MouseOut{
        $sb -> Text('');
}

sub Open1_Click{

        my @parms;
        push @parms,
          -filter =>
                [ 'TXT - Tab�ָ��ı�', '*.txt'
                ],
          -directory => $ConfigHash{InputLoc},
          -title => 'ѡ���ļ�',
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

sub Open1_MouseMove{
        $sb -> Text('������λ�ö�ȡ');
}

sub Open1_MouseOut{
        $sb -> Text('');
}

sub List2_DblClick{
        if (%PrevAllele){
                $error = "�Ѿ��ɹ���ȡ���ݣ��Ƿ����¶�ȡ��";
                my $s = Win32::MsgBox $error,1, "ע�⣡";
                return 0 if $s != 1;
        }

        $SummaryLoaded = 0;
        $display2->Enable($InputLoaded*$SummaryLoaded);
        $open3->Enable($InputLoaded*$SummaryLoaded);

        $error = "��δ��ȡ";
        $text23 -> Text($error);

        my $sel = $Input2->GetCurSel();
        $PrevExp_str = $PrevFound[$sel];

        %PrevAllele = ();

        unless ($InputLoaded){
                $error = "���ȶ�ȡ������Ϣ��\n";
                Win32::MsgBox $error, 0, "����";
                return 0;
        }

        unless (open IN,$PrevExp_str){
                $error = "�ļ���ʧ�ܣ�\n";
                Win32::MsgBox $error, 0, "����";
                return 0;
        }
        my $tmp = <IN>;
        chomp $tmp;
        my $yes = 1;
        my @tmp = split /\t/, $tmp;
        $yes = 0 if @tmp != 17;
        # print "YES:",$yes,"\n";
        foreach $i(1..16){
                $yes = 0 unless exists $markerExist{$tmp[$i]};
                # print $tmp[$i],"|",$yes,"\n";
        }

        if ($yes != 1){
                $error = "����ļ�ò�Ʋ���\n";
                Win32::MsgBox $error, 0, "����";
                return;
        }
        print $PrevAllele{STR154465}{D7S820},"\n";
        while (<IN>){
                chomp;
                my @str = split "\t", $_;
                unless (exists $exp_id{$str[0]}){
                        # print "Next!\n" if $str[0] eq 'STR154465';
                        next;
                }
                ###����Ҫ###
                if (exists $HasChimerism{$str[0]}){
                        print $HasChimerism{$str[0]},"\n";
                        next;
                }
                ###�˴������д���ᵼ���µ�Allele�޷���ȡ###
                my $num = shift @str;
                foreach my $tmp(@markers){
                        $PrevAllele{$num}{$tmp} = shift @str;
                        # print $num,"|",$tmp,"|", $PrevAllele{$num}{$tmp} ,"\n" if $PrevAllele{$num}{$tmp}=~ /\s/;
                        $PrevAllele{$num}{$tmp} =~ s/\s//g;
                }
        }
        close IN;
        # print $PrevAllele{STR154465}{D7S820},"\n";
        $error = "��ȡ�ɹ���";
        $text23 -> Text($error);
        $curr_index = 0;

        $SummaryLoaded = 1;
        $display2->Enable($InputLoaded*$SummaryLoaded);
        $open3->Enable($InputLoaded*$SummaryLoaded);
}

sub List2_MouseMove{
        $sb -> Text('˫����Ŀ�Զ�ȡ');
}

sub List2_MouseOut{
        $sb -> Text('');
}

sub Open2_Click{
        my @parms;
        push @parms,
          -filter =>
                [ 'TXT - Tab�ָ��ı�', '*.txt'
                ],
          -directory => $ConfigHash{SummaryLoc},
          -title => 'ѡ���ļ�',
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

sub Open2_MouseMove{
        $sb -> Text('������λ�ö�ȡ');
}

sub Open2_MouseOut{
        $sb -> Text('');
}

sub DISPLAY2_Click{
        unless (%PrevAllele){
                $error = "��δ��ȡ�ļ���\n";
                Win32::MsgBox $error, 0, "����";
                return 0;
        }

A:
        my $Temp_typinglist = sprintf "�����б�-%4d%02d%02d.xlsx",$year, $mon, $mday;
        my $workbook;
        unless ($workbook = Excel::Writer::XLSX->new($Temp_typinglist)){
                $error = $Temp_typinglist."����ʹ���У�
��رպ����ԣ�";
                Win32::MsgBox $error, 0, "����";
                return 0;
        }

        my $format1 = $workbook->add_format(
                        size            => 9,
                        bold            => 0,
                        align           => 'left',
                        font            => decode('GB2312','����'),
                        'top'           => 1,
                        'bottom'        => 1,
                        'left'          => 1,
                        'right'         => 1,
        );

        my $worksheet = $workbook->add_worksheet();
        $worksheet->hide_gridlines();
        $worksheet->keep_leading_zeros();
        $worksheet->set_landscape();
        $worksheet->set_paper(9);
        $worksheet->set_margin_left(0.394);
        $worksheet->set_margin_right(0.394);
        $worksheet->set_column(0,2, 10);
        $worksheet->set_column(3,3, 1.75);
        $worksheet->set_column(4,5, 10);
        $worksheet->set_column(6,7, 1.75);
        $worksheet->set_column(7,8, 10);
        $worksheet->set_column(9,9, 1.75);
        $worksheet->set_column(10,11, 10);
        $worksheet->set_column(12,12, 1.75);
        $worksheet->set_column(13,15, 10);

        my $pages = int(($#TCA_id+1) / 10)+1;

        foreach my $i(0..$pages*39-1){
                $worksheet->set_row($i, 12.7);
        }

        foreach my $i(1..$pages){
                # $worksheet->write(($i-1)*38,0,' ', $format1);
                # $worksheet->write(($i-1)*38+1,0,' ', $format1);
                # $worksheet->write(($i-1)*38+19,0,' ', $format1);
                # $worksheet->write(($i-1)*38+20,0,' ', $format1);
                my $j = 2;
                foreach (@markers){
                        $worksheet->write(($i-1)*38+$j,0,$markers[$j-2], $format1);
                        $worksheet->write(($i-1)*38+$j,15,$markers[$j-2], $format1);
                        $worksheet->write(($i-1)*38+$j+19,0,$markers[$j-2], $format1);
                        $worksheet->write(($i-1)*38+$j+19,15,$markers[$j-2], $format1);
                        $j ++;
                }
        }

        foreach my $i(0.. $#TCA_id){
                my $TCAID = $TCA_id[$i];
                my @seq = split ",", $exp_seq{$TCAID};
                my $AAA = $exp_list[$seq[0]];
                my $BBB = $exp_list[$seq[1]];
                my $j = 2;
                my $strA;
                my $strB;
                $worksheet->write(int($i/5)*19,$i%5*3+1,$data_in[$seq[-1]][7], $format1);
                $worksheet->write(int($i/5)*19,$i%5*3+2,decode('GB2312', $data_in[$seq[-1]][9]), $format1);
                $worksheet->write(int($i/5)*19+1,$i%5*3+1,$AAA, $format1);
                $worksheet->write(int($i/5)*19+1,$i%5*3+2,$BBB, $format1);
                foreach (@markers){
                        unless (exists $PrevAllele{$AAA}){
                                $strA = ' ';
                        }else{
                                $strA =  $PrevAllele{$AAA}{$_};
                        }
                        unless (exists $PrevAllele{$BBB}){
                                $strB = ' ';
                        }else{
                                $strB =  $PrevAllele{$BBB}{$_};
                        }
                        $worksheet->write(int($i/5)*19+$j,$i%5*3+1, decode('GB2312', $strA), $format1);
                        $worksheet->write(int($i/5)*19+$j,$i%5*3+2, decode('GB2312', $strB), $format1);
                        $j ++;
                }
        }


        $workbook -> close();
        `start $Temp_typinglist`;

        return 0;
}

sub DISPLAY2_MouseMove{
        $sb -> Text('�ο����з��ͽ����ӡ����ʵ������');
}

sub DISPLAY2_MouseOut{
        $sb -> Text('');
}

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

sub Open3_Click{
        unless (@exp_list){
                $error = "���ȶ�ȡ������Ϣ��\n";
                Win32::MsgBox $error, 0, "����";
                return 0;
        }

        unless (%PrevAllele){
                $error = "���ȶ�ȡ���з�����Ϣ��\n";
                Win32::MsgBox $error, 0, "����";
                return 0;
        }

        my @parms;
        push @parms,
          -multisel => 20,
          -filter =>
                [ 'TXT - Tab�ָ��ı�', '*.txt'
                ],
          -directory => $ConfigHash{ThisLoc},
          -title => 'ѡ���ļ�',
          -parent => $main,
          -owner => $main;
        my @file = Win32::GUI::GetOpenFileName ( @parms );
        # print "$_\n" for @file;
        return 0 unless $file[0];
        if (@file == 1){
                chomp $file[0];
                push @ThisFound, $file[0];
                push @ThisList, &Shorten($file[0], 57);
                $Input3 -> Enable(1);
                $Input3 -> Add(&Shorten($file[0], 57));
                $Read3->Enable(1);
                return 0;
        }
        ##�����ѡ�����ظ�ʽΪ ·��;�ļ���1;�ļ���2...
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

sub Open3_MouseMove{
        $sb -> Text('ѡ���»������ļ�����ӵ��Ҳ��б���');
}

sub Open3_MouseOut{
        $sb -> Text('');
}

sub List3_SelChange{
        my @sel = $Input3->GetSelItems();
        if (@sel > 0){
                $del3 -> Enable(1);
        }else{
                $del3 -> Enable(0);
        }
}

sub List3_MouseMove{
        $sb -> Text('ѡ����Ŀ���и������(֧��Ctrl��Shift���ж�ѡ)');
}

sub List3_MouseOut{
        $sb -> Text('');
}

sub Read3_Click{
        if (%ThisAllele){
                $error = "�Ѿ��ɹ���ȡ���ݣ��Ƿ����¶�ȡ��";
                my $s = Win32::MsgBox $error,1, "ע�⣡";
                return 0 if $s != 1;
        }

        %ThisAllele = ();
        $ExpLoaded = 0;
        $run4 -> Enable(0);
        $text3 -> Text("��δ��ȡ");

        foreach my $file (@ThisFound){
                next if $file =~/^\./;
                if (open IN,$file){

                }else{
                        $error = $file."��ʧ�ܣ�\n";
                        Win32::MsgBox $error, 0, "����";
                        return 0;
                };
                my %USE = ();
                while(<IN>){
                        chomp;
                        next if /Sample\sName/;
                        next if /LADDER/;
                        next if /NC\s/;
                        # next if /QC\d+\s/;

                        my @line = split /\t/,$_;
                        my ($tmpallele, $tmparea, $num);
                        if ($line[2] =~ /vWA/){$line[2] =~ s/vWA/VWA/;}
                        if ($line[2] =~ /AMEL/){$line[2] =~ s/AMEL/Amel/;}
                        my $found = 0;
                        if (exists $exp_id{$line[0]}){
                                $num = $line[0];
                                $found = 1;
                                # print "$num �ҵ��ˣ�\n";
                        }elsif($line[0] =~ /^(TB\d+)/){
                                my $tmpstr = $1;

                                if (exists $trans{$tmpstr}){
                                        if ($trans{$tmpstr} eq "ERROR"){
                                                $num = $line[0];
                                        }else{
                                                $num = $trans{$tmpstr};
                                                $found = 1;
                                        }
                                }else{
                                        foreach my $str(@exp_list){
                                                if ($str =~ /$tmpstr$/i){
                                                        $found = 1;
                                                        $num = $str;
                                                        $trans{$tmpstr} = $str;
                                                        # print "$tmpstr --> $str\n";
                                                        last;
                                                }
                                        }
                                        if ($found == 0){
                                                # print "δ�ҵ�",$line[0],"��ʵ���¼��\n";
                                                $trans{$tmpstr} = "ERROR";
                                                $num = $line[0];
                                        }
                                }
                        }elsif($line[0] =~ /^(\d{3,7})-?([A-Z]*)$/){
                                my $tmpstr = $2 ? $1.'-'.$2 : $1;

                                if (exists $trans{$tmpstr}){
                                        if ($trans{$tmpstr} eq "ERROR"){
                                                $num = $line[0];
                                        }else{
                                                $num = $trans{$tmpstr};
                                                $found = 1;
                                        }
                                }else{
                                        foreach my $str(@exp_list){
                                                if ($str =~ /$tmpstr$/i){
                                                        $found = 1;
                                                        $num = $str;
                                                        $trans{$tmpstr} = $str;
                                                        # print "$tmpstr --> $str\n";
                                                        last;
                                                }
                                        }
                                        if ($found == 0){
                                                # print "δ�ҵ�",$line[0],"��ʵ���¼��\n";
                                                $trans{$tmpstr} = "ERROR";
                                                $num = $line[0];
                                        }
                                }
                        }else{
                                # print "ʵ�����",$line[0],"�д������飡\n";
                                next;
                        }

                        next if $found == 0;
                        #print $file,"|",$num,"\n";
                        if    ($line[6]){$tmpallele = join "��", ($line[3],$line[4],$line[5],$line[6]);}
                        elsif ($line[5]){$tmpallele = join "��", ($line[3],$line[4],$line[5]);}
                        elsif ($line[4]){$tmpallele = join "��", ($line[3],$line[4]);}
                        else            {$tmpallele =             $line[3];}

                        if    ($line[10]){$tmparea = join "��", ($line[7],$line[8],$line[9],$line[10]);}
                        elsif ($line[9]) {$tmparea = join "��", ($line[7],$line[8],$line[9]);}
                        elsif ($line[8]) {$tmparea = join "��", ($line[7],$line[8]);}
                        else             {$tmparea =             $line[7];}

                        # if (exists $PrevAllele{$num}{$line[2]}){
                                # $ThisAllele{$num}{$line[2]} = $tmpallele;
                                # $area  {$num}{$line[2]} = $tmparea;
                        # }
                        #
                        if (exists $PrevAllele{$num}){
                                print "$num ���У�\n";
                                if (exists $USE{$num}){
                                        next;
                                }else{
                                        $error = "������� ".$num." ���з������ݣ������»������Ƿ�ʹ�ã�
����������ǻ������󣬱������ݲ�ʹ�ý��ᵼ�´���";
                                        my $s = Win32::MsgBox $error,4, "ע�⣡";
                                        if ($s == 6){
                                                delete $PrevAllele{$num};
                                                $ThisAllele{$num}{$line[2]} = $tmpallele;
                                                $area{$num}{$line[2]} = $tmparea;
                                        }else{
                                                $USE{$num} = 'no';
                                                next;
                                        }
                                }
                        }else{
                                $ThisAllele{$num}{$line[2]} = $tmpallele;
                                $area{$num}{$line[2]} = $tmparea;
                        }
                        #


                        #print $file,"|",$num,"|",$line[2],"|",$allele{$num}{$line[2]},"|",$area{$num}{$line[2]},"\n";

                }
                close IN;

        }
        $error = "��ȡ�ɹ�";
        $text3 -> Text($error);
        $ExpLoaded = 1;
        $run4 -> Enable(1);
        return 0;
}

sub Read3_MouseMove{
        $sb -> Text('��ȡ����б��е�����');
}

sub Read3_MouseOut{
        $sb -> Text('');
}

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
                $text3 -> Text("��δ��ȡ");
                $Read3 -> Enable(0);
                $Input3 -> Enable(0);
                $del3 -> Enable(0);
        }
}

sub Del3_MouseMove{
        $sb -> Text('�Ƴ��Ҳ��б���ѡ�е��ļ�');
}

sub Del3_MouseOut{
        $sb -> Text('');
}

sub RUN_Click{

        # Righto... I found some stuff, but not exactly what I was after
# From the Documentation...
# BrowseForFolder( OPTIONS ) Displays the standard ``Browse For Folder'' dialog box. Returns the selected item's name, or undef if no item was selected or an error occurred. Note that BrowseForFolder must be called as a standalone function, not as a method. Example:
        # $folder = Win32::GUI::BrowseForFolder(
                # -root => "C:\\Program Files",
                # -includefiles => 1,
        # );
#
# other options are... -computeronly, -domainonly, -driveonly, -editbox, -folderonly, -includefiles, -owner, -printeronly, -root, -title
# What I was actually thinking of was
# $ret = GUI::GetSaveFileName(
    # -title  => "Save your newly generated Mail Merge Document.",
    # -file   => "\0" . " " x 256,
    # -filter => [
        # "Word documents (*.doc)" => "*.doc",
        # "All files", "*.*",
    # ],
# With another option in there, but I can't find it anywhere. Oh well, I am sure that browse for folder will work. Let me know how you get on.
# Regards, Gerard.
#
# https://metacpan.org/pod/distribution/Win32-GUI/GENERATED/Win32/GUI/Reference/Methods.pod#BrowseForFolder
#
# -title => STRING
    # the title for the dialog
# -computeronly => 0/1 (default 0)
    # only enable computers to be selected
# -domainonly => 0/1 (default 0)
    # only enable computers in the current domain or workgroup
# -driveonly => 0/1 (default 0)
    # only enable drives to be selected
# -editbox => 0/1 (default 0)
    # if 1, the dialog will include an edit field in which
    # the user can type the name of an item
# -folderonly => 0/1 (default 0)
    # only enable folders to be selected
# -includefiles => 0/1 (default 0)
    # the list will include files as well folders
# -newui => 0/1 (default 0)
    # use the "new" user interface (which has a "New folder" button)
# -nonewfolder => 0/1 (default 0)
    # hides the "New folder" button (only meaningful with -newui => 1)
# -owner => WINDOW
    # A Win32::GUI::Window or Win32::GUI::DialogBox object specifiying the
    # owner window for the dialog box
# -printeronly => 0/1 (default 0)
    # only enable printers to be selected
# -directory => PATH
    # the default start directory for browsing
# -root => PATH or CONSTANT
    # the root directory for browsing; this can be either a
    # path or one of the following constants (minimum operating systems or
    # Internet Explorer versions that support the constant are shown in
    # square brackets. NT denotes Windows NT 4.0, Windows 2000, XP, etc.):
        #
        # CSIDL_FLAG_CREATE (0x8000)
         # [2000/ME] Combining this with any of the constants below will create the folder if it does not already exist.
     # CSIDL_ADMINTOOLS (0x0030)
         # [2000/ME] Administrative Tools directory for current user
     # CSIDL_ALTSTARTUP (0x001d)
         # [All] Non-localized Startup directory in the Start menu for current user
     # CSIDL_APPDATA (0x001a)
         # [IE4] Application data directory for current user
     # CSIDL_BITBUCKET (0x000a)
         # [All] Recycle Bin
     # CSIDL_CDBURN_AREA (0x003b)
         # [XP] Windows XP directory for files that will be burned to CD
     # CSIDL_COMMON_ADMINTOOLS (0x002f)
         # [2000/ME] Administrative Tools directory for all users
     # CSIDL_COMMON_ALTSTARTUP (0x001e)
         # [All] Non-localized Startup directory in the Start menu for all users
     # CSIDL_COMMON_APPDATA (0x0023)
         # [2000/ME] Application data directory for all users
     # CSIDL_COMMON_DESKTOPDIRECTORY (0x0019)
         # [NT] Desktop directory for all users
     # CSIDL_COMMON_DOCUMENTS (0x002e)
         # [IE4] My Documents directory for all users
     # CSIDL_COMMON_FAVORITES (0x001f)
         # [NT] Favorites directory for all users
     # CSIDL_COMMON_MUSIC (0x0035)
         # [XP] Music directory for all users
     # CSIDL_COMMON_PICTURES (0x0036)
         # [XP] Image directory for all users
     # CSIDL_COMMON_PROGRAMS (0x0017)
         # [NT] Start menu "Programs" directory for all users
     # CSIDL_COMMON_STARTMENU (0x0016)
         # [NT] Start menu root directory for all users
     # CSIDL_COMMON_STARTUP (0x0018)
         # [NT] Start menu Startup directory for all users
     # CSIDL_COMMON_TEMPLATES (0x002d)
         # [NT] Document templates directory for all users
     # CSIDL_COMMON_VIDEO (0x0037)
         # [XP] Video directory for all users
     # CSIDL_CONTROLS (0x0003)
         # [All] Control Panel applets
     # CSIDL_COOKIES (0x0021)
         # [All] Cookies directory
     # CSIDL_DESKTOP (0x0000)
         # [All] Namespace root (shown as "Desktop", but is parent to my computer, control panel, my documents, etc.)
     # CSIDL_DESKTOPDIRECTORY (0x0010)
         # [All] Desktop directory (for desktop icons, folders, etc.) for the current user
     # CSIDL_DRIVES (0x0011)
         # [All] My Computer (drives and mapped network drives)
     # CSIDL_FAVORITES (0x0006)
         # [All] Favorites directory for the current user
     # CSIDL_FONTS (0x0014)
         # [All] Fonts directory
     # CSIDL_HISTORY (0x0022)
         # [All] Internet Explorer history items for the current user
     # CSIDL_INTERNET (0x0001)
         # [All] Internet root
     # CSIDL_INTERNET_CACHE (0x0020)
         # [IE4] Temporary Internet Files directory for the current user
     # CSIDL_LOCAL_APPDATA (0x001c)
         # [2000/ME] Local (non-roaming) application data directory for the current user
     # CSIDL_MYMUSIC (0x000d)
         # [All] My Music directory for the current user
     # CSIDL_MYPICTURES (0x0027)
         # [2000/ME] Image directory for the current user
     # CSIDL_MYVIDEO (0x000e)
         # [XP] Video directory for the current user
     # CSIDL_NETHOOD (0x0013)
         # [All] My Network Places directory for the current user
     # CSIDL_NETWORK (0x0012)
         # [All] Root of network namespace (Network Neighbourhood)
     # CSIDL_PERSONAL (0x0005)
         # [All] My Documents directory for the current user
     # CSIDL_PRINTERS (0x0004)
         # [All] List of installed printers
     # CSIDL_PRINTHOOD (0x001b)
         # [All] Network printers directory for the current user
     # CSIDL_PROFILE (0x0028)
         # [2000/ME] The current user's profile directory
     # CSIDL_PROFILES (0x003e)
         # [XP] The directory that holds user profiles (see CSDIL_PROFILE)
     # CSIDL_PROGRAM_FILES (0x0026)
         # [2000/ME] Program Files directory
     # CSIDL_PROGRAM_FILES_COMMON (0x002b)
         # [2000] Directory for files that are used by several applications. Usually Program Files\Common
     # CSIDL_PROGRAMS (0x0002)
         # [All] Start menu "Programs" directory for the current user
     # CSIDL_RECENT (0x0008)
         # [All] Recent Documents directory for the current user
     # CSIDL_SENDTO (0x0009)
         # [All] "Send To" directory for the current user
     # CSIDL_STARTMENU (0x000b)
         # [All] Start Menu root for the current user
     # CSIDL_STARTUP (0x0007)
         # [All] Start Menu "Startup" folder for the current user
     # CSIDL_SYSTEM (0x0025)
         # [2000/ME] System directory. Usually \Windows\System32
     # CSIDL_TEMPLATES (0x0015)
         # [All] Document templates directory for the current user
     # CSIDL_WINDOWS (0x0024)
         # [2000/ME] Windows root directory, can also be accessed via the environment variables %windir% or %SYSTEMROOT%.

        my $ret = Win32::GUI::BrowseForFolder (
                -title      => "��ѡ�񱣴�·��",
                # -editbox    => 1,
                -directory  => $ConfigHash{OutputLoc},
                -folderonly => 1,
                -newui      => 1,
                -parent => $main,
                -owner => $main,
        );
        return 0 unless $ret;
        $Output_Dir = $ret;
        $ConfigHash{OutputLoc} = $ret;

        $sb->Move( 0, ($main->ScaleHeight() - $sb->Height()) );
        $sb->Resize( $main->ScaleWidth(), $sb->Height() );
        $sb->Text("���ںϲ������ļ�...");
        $RUNwindow -> Show();

        %allele = ();

        foreach my $PrevKey1(keys %PrevAllele){
                foreach my $PrevKey2(keys %{$PrevAllele{$PrevKey1}}){
                        $allele{$PrevKey1}{$PrevKey2} = $PrevAllele{$PrevKey1}{$PrevKey2};
                        # print "Prev $PrevKey1|$PrevKey2|",$allele{$PrevKey1}{$PrevKey2},"\n";
                }
        }
        foreach my $ThisKey1(keys %ThisAllele){
                foreach my $ThisKey2(keys %{$ThisAllele{$ThisKey1}}){
                        $allele{$ThisKey1}{$ThisKey2} = $ThisAllele{$ThisKey1}{$ThisKey2};
                        # print "This $ThisKey1|$ThisKey2|",$allele{$ThisKey1}{$ThisKey2},"\n";
                }
        }

        my (%date4,%date1,%date2,%sample,%operation,%cells,%date3,%number,%rptnum,%name,%patient,%gender,%age,%diagnosis,%relation,%xnum,%hospital,%doctor,%hosptl_num,%bed_num);
        my %sheet_name;

        foreach (keys %exp_id){

                my $number = $exp_id{$_};

                $date4{$_}     = $data_in[$number][0];       #��������
                $date1{$_}     = $data_in[$number][1];       #����ʱ��
                $date2{$_}     = $data_in[$number][2];       #��ֲ����
                $sample{$_}    = $data_in[$number][3];
                $operation{$_} = $data_in[$number][4];
                $cells{$_}     = $data_in[$number][5];
                $date3{$_}     = $data_in[$number][6];       #��������
                $number{$_}    = $data_in[$number][7];
                $rptnum{$_}    = $data_in[$number][8];
                $name{$_}      = $data_in[$number][9];
                $patient{$_}   = $data_in[$number][10];
                $gender{$_}    = $data_in[$number][11];
                $age{$_}       = $data_in[$number][12];
                $diagnosis{$_} = $data_in[$number][13];
                $relation{$_}  = $data_in[$number][14];
                $xnum{$_}      = $data_in[$number][15];
                $hospital{$_}  = $data_in[$number][17];
                $doctor{$_}    = $data_in[$number][18];
                $hosptl_num{$_}= $data_in[$number][19];
                $bed_num{$_}   = $data_in[$number][20];
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
        foreach (@markers){$allele{'  '}{$_} = ' '; }
        foreach (@markers){$area  {'  '}{$_} = ' '; }


        $sb->Move( 0, ($main->ScaleHeight() - $sb->Height()) );
        $sb->Resize( $main->ScaleWidth(), $sb->Height() );
        $sb->Text("���������ļ�...");

        my $success = 1;


        my @conclusion;
        my @num1;
        my @num2;
        my @num3;
        my @sheet;

        # my @count_sum;
        my @count_n;
        my @count_avg;
        my @SD;
        my @marker_type;
        my @type;
        my @count;
        foreach my $z(0 .. $#TCA_id){
                my $TCAID = $TCA_id[$z];
                if ($exp_error{$TCAID} == 1){
                        $conclusion[$z] = '����';
                        next;
                }
                # print STDERR $TCAID,"ʵ������",$exp_num{$TCAID},"\n";
                $RptBox -> Append('׼��'.$TCAID.'...');
                if ($exp_num{$TCAID} == 1){
                        $conclusion[$z] = '��';
                        my @seq = split ",", $exp_seq{$TCAID};
                        unless (exists $allele{$exp_list[$seq[0]]}){
                                $error = '�»�������δ�ҵ����'.$exp_list[$seq[0]].'������
���飡';
                                Win32::MsgBox $error, 0, "����";
                                $success = 0;
                                $RptBox -> Append("ʧ�ܡ��»����ݲ�ȫ��\r\n");
                                $conclusion[$z] = '����';
                                next;
                        }
                        if ($data_in[$seq[0]][4] eq "��ǰ"){
                                if ($data_in[$seq[0]][10] eq "����"){
                                        $num1[$z] = $exp_list[$seq[0]];
                                        $num2[$z] = '  ';
                                        $num3[$z] = '  ';
                                        $sheet[$z] = $name{$num1[$z]};
                                }else{
                                        $num1[$z] = '  ';
                                        $num2[$z] = $exp_list[$seq[0]];
                                        $num3[$z] = '  ';
                                        $sheet[$z] = $name{$num2[$z]};
                                }
                        }else{
                                if ($data_in[$seq[0]][10] eq "����"){
                                        $num1[$z] = '  ';
                                        $num2[$z] = '  ';
                                        $num3[$z] = $exp_list[$seq[0]];
                                        $sheet[$z] = $name{$num3[$z]};
                                }else{
                                        my $error  = "�����ţ�".$TCAID."ֻ����һ����ʾΪ���������������\n���飬���ν������ɱ��棡\n";
                                        Win32::MsgBox $error, 0, "ע�⣡";
                                        $RptBox -> Append("ʧ�ܡ������ߡ�\r\n");
                                        $conclusion[$z] = '����';
                                        next;
                                }
                        }
                        $RptBox -> Append("�ɹ���\r\n");
                        # printf "%s|%s|%s|%s|%s\n",        $TCAID, $num1[$z], $num2[$z], $num3[$z], $sheet[$z];
                }elsif ($exp_num{$TCAID} == 2){
                        $conclusion[$z] = '��';
                        my @seq = split ",", $exp_seq{$TCAID};
                        unless (exists $allele{$exp_list[$seq[0]]}){
                                $error = '�»�������δ�ҵ����'.$exp_list[$seq[0]].'������
���飡';
                                Win32::MsgBox $error, 0, "����";
                                $success = 0;
                                $RptBox -> Append("ʧ�ܡ��»����ݲ�ȫ��\r\n");
                                $conclusion[$z] = '����';
                                next;
                        }
                        unless (exists $allele{$exp_list[$seq[1]]}){
                                $error = '�»�������δ�ҵ����'.$exp_list[$seq[1]].'������
���飡';
                                Win32::MsgBox $error, 0, "����";
                                $success = 0;
                                $RptBox -> Append("ʧ�ܡ��»����ݲ�ȫ��\r\n");
                                $conclusion[$z] = '����';
                                next;
                        }
                        my $sum = 0;
                        foreach my $i(@seq){
                                if ($data_in[$i][4] eq "��ǰ"){
                                        $sum += 0;
                                }else{
                                        $sum += 2;
                                }
                                if ($data_in[$i][10] eq "����"){
                                        $sum += 4;
                                }else{
                                        $sum += 8;
                                }
                        }
                        if ($sum == 12){
                                $num1[$z] = $exp_list[$seq[0]];
                                $num2[$z] = $exp_list[$seq[1]];
                                $num3[$z] = '  ';
                                $sheet[$z] = $name{$num1[$z]};
                        }elsif($sum == 10){
                                $num1[$z] = $exp_list[$seq[0]];
                                $num2[$z] = '  ';
                                $num3[$z] = $exp_list[$seq[1]];
                                $sheet[$z] = $name{$num1[$z]};
                        }elsif($sum == 14){
                                $num1[$z] = '  ';
                                $num2[$z] = $exp_list[$seq[0]];
                                $num3[$z] = $exp_list[$seq[1]];
                                $sheet[$z] = $name{$num3[$z]};
                        }else{
                                my $error  = "�����ţ�".$TCAID."����һ����ʾΪ���������������
���飬���ν������ɱ��棡";
                                Win32::MsgBox $error, 0, "ע�⣡";
                                $RptBox -> Append("ʧ�ܡ������ߡ�\r\n");
                                $conclusion[$z] = '����';
                                next;
                        }
                        $RptBox -> Append("�ɹ���\r\n");
                        ####
                        #����ͨ�����������ж�2������������������߻��ǹ���
                        ####
                        # printf "%s|%s|%s|%s|%s\n",        $TCAID, $num1[$z], $num2[$z], $num3[$z],$sheet[$z];
                }elsif ($exp_num{$TCAID} == 3){
                        my @seq = split ",", $exp_seq{$TCAID};
                        # print $exp_list[$seq[0]],"|", $allele{$exp_list[$seq[0]]}{D7S820},"\n";
                        # print $exp_list[$seq[1]],"|", $allele{$exp_list[$seq[1]]}{D7S820},"\n";
                        # print $exp_list[$seq[2]],"|", $allele{$exp_list[$seq[2]]}{D7S820},"\n";

                        unless (exists $allele{$exp_list[$seq[0]]}){
                                $error = '�»�������δ�ҵ����'.$exp_list[$seq[0]].'������
���飡';
                                Win32::MsgBox $error, 0, "����";
                                $success = 0;
                                $RptBox -> Append("ʧ�ܡ��»����ݲ�ȫ��\r\n");
                                $conclusion[$z] = '����';
                                next;
                        }
                        unless (exists $allele{$exp_list[$seq[1]]}){
                                $error = '�»�������δ�ҵ����'.$exp_list[$seq[1]].'������
���飡';
                                Win32::MsgBox $error, 0, "����";
                                $success = 0;
                                $RptBox -> Append("ʧ�ܡ��»����ݲ�ȫ��\r\n");
                                $conclusion[$z] = '����';
                                next;
                        }
                        unless (exists $allele{$exp_list[$seq[2]]}){
                                $error = '�»�������δ�ҵ����'.$exp_list[$seq[2]].'������
���飡';
                                Win32::MsgBox $error, 0, "����";
                                $success = 0;
                                $RptBox -> Append("ʧ�ܡ��»����ݲ�ȫ��\r\n");
                                $conclusion[$z] = '����';
                                next;
                        }
                        $num1[$z] = $exp_list[$seq[0]];
                        $num2[$z] = $exp_list[$seq[1]];
                        $num3[$z] = $exp_list[$seq[2]];
                        $sheet[$z] = $name{$num1[$z]};
                        # printf "%s|%s|%s|%s|%s\n",        $TCAID, $num1, $num2, $num3,$sheet;
                }else{
                        my $error  = "�����ţ�".$TCAID."��ʵ��������Ź��࣡���飡\n";
                        Win32::MsgBox $error, 0, "ע�⣡";
                        $RptBox -> Append("ʧ�ܡ��������ࡿ\r\n");
                        $conclusion[$z] = '����';
                        next;
                }
                # print $num2,"|",$relation{$num2},"\n";
                # $count_sum[$z] = 0;
                $count_n[$z] = 0;
                $count_avg[$z] = 0;
                $SD[$z] = 0;

                my $errorcount = 0;
                foreach my $k (0..$#markers){

                        # if ($conclusion[$z]){
                                # if ($conclusion[$z] eq '����'){
                                        # last;
                                # }
                        # }
                        my %alleles_before = ();

                        # print $num1[$z],"|",$allele{$num1[$z]}{$markers[$k]},"\n";
                        # print $num2[$z],"|",$allele{$num2[$z]}{$markers[$k]},"\n";
                        # print $num3[$z],"|",$allele{$num3[$z]}{$markers[$k]},"\n";
                        # print $num3[$z],"|",$area{$num3[$z]}{$markers[$k]},"\n";

                        my @allele1 = split/��/, $allele{$num1[$z]}{$markers[$k]};
                        $alleles_before{$_} = 1 foreach @allele1;
                        my @allele2 = split/��/, $allele{$num2[$z]}{$markers[$k]};
                        $alleles_before{$_} = 1 foreach @allele2;
                        my @allele3 = split/��/, $allele{$num3[$z]}{$markers[$k]};

                        foreach (@allele3){
                                if (!exists $alleles_before{$_}){
                                        $type[$z][$k] = "error";
                                        $count[$z][$k] = "error";
                                        $errorcount ++;
                                        last;
                                }
                        }

                        # foreach (@allele3){
                                # if (!exists $alleles_before{$_}){
                                        # $success = 0;
                                        # $error = '���浥��'.$TCAID.'��Ӧ����ǰ������ʹ������飡
# '.$markers[$k].':
# ��ǰ���ߣ�'.$num1[$z].'|'.$allele{$num1[$z]}{$markers[$k]}.'
# ���ߣ�'.$num2[$z].'|'.$allele{$num2[$z]}{$markers[$k]}.'
# �����ߣ�'.$num3[$z].'|'.$allele{$num3[$z]}{$markers[$k]}.'
# ���������ߴ˷ݱ��档';
                                        # Win32::MsgBox $error, 0, "ע��";
                                        # $RptBox -> Append("ʧ�ܡ��������ݴ��󡿣�\r\n");
                                        # $RptBox -> Append($markers[$k].":\r\n��ǰ���ߣ�".$num1[$z]."|".$allele{$num1[$z]}{$markers[$k]}."\r\n    ���ߣ�".$num2[$z]."|".$allele{$num2[$z]}{$markers[$k]}."\r\n�����ߣ�".$num3[$z]."|".$allele{$num3[$z]}{$markers[$k]}."\r\n");
                                        # $conclusion[$z] = '����';
                                        # last;
                                # }
                        # }

                        my @area3   = split/��/, $area{$num3[$z]}{$markers[$k]};
                        # print $_,"|" foreach @allele1;
                        # print $_,"|" foreach @allele2;
                        # print $_,"|" foreach @allele3;
                        # print $_,"|" foreach @area3;
                        # print "\n";
                }

                if ($errorcount >= 6){
                        $success = 0;
                        $error = '���浥��'.$TCAID.'��16��λ����
'.$errorcount.'�����ʹ������飡
���������ߴ˷ݱ��档';
                        Win32::MsgBox $error, 0, "ע��";
                        $RptBox -> Append("ʧ�ܡ��������ݴ���\r\n");
                        $conclusion[$z] = '����';
                        next;
                }

                foreach my $k (0..$#markers){
                        next if $count[$z][$k] eq 'error';
                        my @allele1 = split/��/, $allele{$num1[$z]}{$markers[$k]};
                        my @allele2 = split/��/, $allele{$num2[$z]}{$markers[$k]};
                        my @allele3 = split/��/, $allele{$num3[$z]}{$markers[$k]};
                        my @area3   = split/��/, $area{$num3[$z]}{$markers[$k]};

                        if ($allele{$num1[$z]}{$markers[$k]} eq $allele{$num2[$z]}{$markers[$k]}){
                        #��ͬ   (A,A || AB,AB)
                                $type[$z][$k] = '';
                                $count[$z][$k] = '';
                        }elsif ($markers[$k] eq 'Amel'){
                                $type[$z][$k] = '';
                                $count[$z][$k] = '';
                        }elsif (@allele2 == 1 && @allele1 == 2 && ($allele1[0] eq $allele2[0] || $allele1[1] eq $allele2[0])){
                        #���ߴ���&&������һ����ͬ   (AB,A || AB,B)
                                $type[$z][$k] = 2;
                        }elsif (@allele1 == 1 && @allele2 == 2 && ($allele1[0] eq $allele2[0] || $allele1[0] eq $allele2[1])){
                        #���ߴ���&&������һ����ͬ   (A,AB || B,AB)
                                $type[$z][$k] = 3;
                        }elsif ((@allele1 == 2 && @allele2 == 2 && $allele1[0] ne $allele2[0] && $allele1[0] ne $allele2[1]  && $allele1[1] ne $allele2[0] && $allele1[1] ne $allele2[1]) ||@allele1 == 1 && @allele2 == 2 && $allele1[0] ne $allele2[0] && $allele1[0] ne $allele2[1] || @allele1 == 2 && @allele2 == 1 && $allele1[0] ne $allele2[0] && $allele1[1] ne $allele2[0] || @allele1 == 1 && @allele2 == 1 && $allele1[0] ne $allele2[0]){
                        #��ȫ��ͬ   (AB,CD || A,CD || AB,C || A,C )
                                $type[$z][$k] = 1;
                        }elsif (  @allele1 == 2 && @allele2 == 2 && (($allele1[0] eq $allele2[0] && $allele1[1] ne $allele2[1]) || ($allele1[1] eq $allele2[1] && $allele1[0] ne $allele2[0]) || ($allele1[1] eq $allele2[0] && $allele1[0] ne $allele2[1]) || ($allele1[0] eq $allele2[1] && $allele1[1] ne $allele2[0]))){
                        #���Ӻ�&&��һ����ͬ  (5 6,5 7 || 5 6,4 6 || 5 6,6 7 || 5 6,4 5)
                                $type[$z][$k] = 4;
                        }else{
                                $type[$z][$k] = "error";
                                $count[$z][$k] = "error";
                        }

                        # print "Type: ",$type[$z][$k],"\n";

                        my %areas;
                        for my $p (0..$#allele3){
                                $areas{$allele3[$p]} = $area3[$p];
                        }
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

                        }elsif ($type[$z][$k] eq 4){
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
                if ($conclusion[$z]){
                        if ($conclusion[$z] eq '����'){
                                next;
                        }
                }
                my @temp_marker = ();
                my @tempcount = ();
                foreach my $k (0..$#markers){
                        if ($count[$z][$k] =~ /\d/){
                                # $count_sum[$z] += $count[$z][$k];
                                # $count_n[$z] += 1;
                                push @tempcount, $count[$z][$k];
                                if ($count[$z][$k]<1 && $count[$z][$k]>0){
                                        $temp_marker[$k] = '���Ƕ��';
                                }else{
                                        $temp_marker[$k] = ' ';
                                }
                        }else{
                                $temp_marker[$k] = ' ';
                        }
                }

                foreach my $k (0..$#markers){
                        $marker_type[$z][$k] = $temp_marker[$k];
                }
                $count_n[$z] = scalar(@tempcount);
                if ($count_n[$z] > 0){
                        ($count_avg[$z], $SD[$z]) = &Avg_SD(@tempcount);
                        $count_avg[$z] = sprintf("%.4f", $count_avg[$z]);
                }else{
                        $success = 0;
                        $error = '���浥��'.$TCAID.'û����Чλ�㣬���飡
���������ߴ˷ݱ��档';
                        Win32::MsgBox $error, 0, "ע��";
                        $RptBox -> Append("ʧ�ܡ�����Чλ�㡿\r\n");
                        $conclusion[$z] = '����';
                        next;
                }

                # if ($count_n[$z] != 0){
                        # $count_avg[$z] = $count_sum[$z] / $count_n[$z];
                        # $count_avg[$z] = sprintf("%.4f", $count_avg[$z]);
                # }

                $RptBox -> Append("�ɹ���\r\n");
                next if $exp_num{$TCAID} != 3;
                ##׷�ӱ���ʵ�������ڴ���
                next if $count_avg[$z] !~ /\d/;
                my $tempid = $identity{$TCAID};
                push @{$Chimerism{$tempid}}, sprintf ("%.2f%s", $count_avg[$z]*100,"%");
                push @{$SampleID{$tempid}}, $num3[$z];
                push @{$ReportDate{$tempid}}, sprintf ("%d-%02d-%02d", $year, $mon, $mday);
                if ($cells{$num3[$z]} ne "-"){
                        $sampleType{$num3[$z]} = $cells{$num3[$z]};
                }else{
                        $sampleType{$num3[$z]} = $sample{$num3[$z]};
                }
                $receiveDate{$num3[$z]} = DateUnify($date1{$num3[$z]});
                $sampleDate{$num3[$z]} = DateUnify($date3{$num3[$z]});
                ##����׷�ӵ��ܱ���
                # print $tempid,"\n";
                # print "Chimerism ";print $_,"|" foreach (@{$Chimerism{$tempid}});print "\n";
                # print "SampleID ";print $_,"|" foreach (@{$SampleID{$tempid}});print "\n";
                # print "ReportDate ";print $_,"|" foreach (@{$ReportDate{$tempid}});print "\n";
                # print "receiveDate ";print $_,"|" foreach (@{$ReportDate{$tempid}});print "\n";
                # print "ReportDate ";print $_,"|" foreach (@{$ReportDate{$tempid}});print "\n";

        }
        my $chimerismSummary = sprintf "Ƕ���ʻ���-%4d%02d%02d.txt",$year, $mon, $mday;
        open SUM,"> $chimerismSummary";
        print SUM "����\tҽԺ\t��������\t�������\t������\tǶ����\t��Чλ��\tSD\tCV\n";

        $RptBox -> Append("���׼����ɣ���ʼ�������\r\n========================\r\n");
        foreach my $z(0..$#TCA_id){
                my $TCAID = $TCA_id[$z];
                $RptBox -> Append($TCAID.'...');
                if ($conclusion[$z] eq '����'){
                        $RptBox -> Append("����\r\n");
                        next;
                }
                if (exists $sheet_name{$sheet[$z]}){
                        $sheet_name{$sheet[$z]} += 1;
                        $Output_rpt_str = sprintf "%s\\%s-%s%s%d.xlsx", $Output_Dir, $TCAID, $sheet[$z], 'AK', $sheet_name{$sheet[$z]};
                }else{
                        $Output_rpt_str = sprintf "%s\\%s-%s%s.xlsx", $Output_Dir, $TCAID, $sheet[$z], 'AK';
                        $sheet_name{$sheet[$z]} = 1;
                }

                my $workbook;
                unless ($workbook = Excel::Writer::XLSX->new($Output_rpt_str)){
                        $error = $Output_rpt_str."
�޷����棡";
                        Win32::MsgBox $error, 0, "����";
                        $success = 0;
                        $RptBox -> Append($Output_rpt_str."��ʧ�ܣ�����\r\n");
                        next;
                }


                my $format1  = $workbook->add_format(size => 18, bold => 1, align => 'center',                      font => decode('GB2312','����')); # HLA�߷ֱ������ͼ�ⱨ��
                my $format2  = $workbook->add_format(size => 11,                                                                                     'top' => 1, 'bottom' => 2);  # ˫��
                my $format3  = $workbook->add_format(size => 11,            align => 'right',  valign => 'vcenter', font => decode('GB2312','����')); # ���浥���
                my $format4  = $workbook->add_format(size => 12, bold => 1, align => 'center', valign => 'vcenter', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # �ͼ쵥λ�������Ŀ write
                my $format5  = $workbook->add_format(size => 12, bold => 1, align => 'center', valign => 'vcenter', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # �ͼ쵥λ�������Ŀ merge
                my $format6  = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # ��Ϣ/��������壬write
                my $format7  = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # ��Ϣ/��������壬merge
                my $format8  = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => 'Times New Roman',       'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # ��Ϣ/�����Times New Roman��write
                my $format9  = $workbook->add_format(size => 11,            align => 'center', valign => 'vcenter', font => 'Times New Roman',       'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # ��Ϣ/�����Times New Roman��merge
                my $format10 = $workbook->add_format(size => 10,            align => 'center', valign => 'vcenter', font => 'Times New Roman',       'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # ��Ϣ/�����Times New Roman��write���ֺ�10
                #my $format11 = $workbook->add_format(size => 10,                               valign => 'vcenter', font => decode('GB2312','����'), text_wrap => 1); # ��ע
                #my $format12 = $workbook->add_format(size => 11, bold => 1,                    valign => 'vcenter', font => decode('GB2312','����')); #
                #my $format13 = $workbook->add_format(size => 11, bold => 1,                    valign => 'vcenter', font => decode('GB2312','����'),             'bottom' => 1);  # �����
                #my $format14 = $workbook->add_format(size => 11, bold => 1,                    valign => 'vcenter', font => 'Times New Roman',                   'bottom' => 1);  # ��������
                my $format15 = $workbook->add_format(size => 12,            align => 'center', valign => 'vcenter', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # 'STRλ��' 'λ��״̬' '��ע'
                my $format16 = $workbook->add_format(size => 9,             align => 'center', valign => 'vcenter', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # ������������
                my $format17 = $workbook->add_format(size => 12, bold => 1,                    valign => 'vcenter', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 0); # '������'
                my $format18 = $workbook->add_format(size => 11,                               valign => 'vcenter', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 0, 'right' => 1); # ������
                my $format19 = $workbook->add_format(size => 8, valign => 'vcenter', font => decode('GB2312','��������'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1, 'text_wrap' => 1,); #��ע

                #####chart�ĸ�ʽ#####
                my $Gfmt1 = $workbook->add_format(size => 10, align => 'right', font => decode('GB2312','����'));  # chart��������
                my $Gfmt2 = $workbook->add_format(size => 14, bold => 1, align => 'center', font => decode('GB2312','����')); # chart����
                my $Gfmt3 = $workbook->add_format(size => 12, bold => 1, align => 'center', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1);# chart��ͷ
                my $Gfmt4 = $workbook->add_format(size => 10, align => 'center', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # chart��
                my $Gfmt5 = $workbook->add_format(size => 14, bold => 1, font => decode('GB2312','����'));  # TCA���ڼ������
                my $Gfmt6 = $workbook->add_format(size => 11, font => decode('GB2312','����'));  # ��ʾ
                my $Gfmt7 = $workbook->add_format(size => 11, align => 'center', font => decode('GB2312','����')); # ��ܰ��ʾ
                my $Gfmt8 = $workbook->add_format(size => 10, align => 'center', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # chart��
                my $Gfmt9 = $workbook->add_format(size =>  9, align => 'center', font => decode('GB2312','����'), 'top' => 1, 'bottom' => 1, 'left' => 1, 'right' => 1); # chart��
                $Gfmt8->set_num_format('0.00');
                #####

                my ($countsheet, $graphic, $worksheet, $graphic_temp);
                $worksheet  = $workbook ->add_worksheet(decode('GB2312',"����"));
                if ($exp_num{$TCAID} == 3 && $sheet_name{$sheet[$z]} == 1){
                        $graphic = $workbook ->add_worksheet(decode('GB2312',"Ƕ������"));
                }
                $countsheet = $workbook->add_worksheet(decode('GB2312',"����"));
                if ($exp_num{$TCAID} == 3 && $sheet_name{$sheet[$z]} == 1){
                        $graphic_temp = $workbook->add_worksheet('temp');
                }

                $countsheet->hide_gridlines();
                $countsheet->keep_leading_zeros();

                my $format101 = $workbook->add_format(size  => 11, font  => decode('GB2312','����'));
                my $format102 = $workbook->add_format(size  => 11, align => 'center', font  => decode('GB2312','����'));
                my $format103 = $workbook->add_format(size  => 11, color => 'red', font   => decode('GB2312','����'));

                $countsheet->write('A01',decode('GB2312','λ��'), $format101);
                $countsheet->merge_range('B1:C1', decode('GB2312','����'), $format102);
                $countsheet->merge_range('D1:E1', decode('GB2312','����'), $format102);
                $countsheet->merge_range('F1:I1', decode('GB2312','����'), $format102);
                $countsheet->write('J01',decode('GB2312','����'), $format101);
                $countsheet->merge_range('K1:N1', decode('GB2312','���'), $format102);
                $countsheet->merge_range('B2:C2', decode('GB2312',$num1[$z]),  $format102);
                $countsheet->merge_range('D2:E2', decode('GB2312',$num2[$z]),  $format102);
                $countsheet->merge_range('F2:I2', decode('GB2312',$num3[$z]),  $format102);
                $countsheet->merge_range('O1:S1', decode('GB2312','Ƕ����'),  $format102);
                $countsheet->write('O2', 'TYPE1',  $format101);
                $countsheet->write('P2', 'TYPE2',  $format101);
                $countsheet->write('Q2', 'TYPE3',  $format101);
                $countsheet->write('R2', 'TYPE4',  $format101);
                $countsheet->write('S2', 'ERROR',  $format101);
                $countsheet->write('T1',decode('GB2312','��Ƕ����'), $format101);

                for my $j (0..$#markers){
                        $countsheet->write($j+2,0,$markers[$j], $format101);
                }
                foreach my $k (0..$#markers){
                        my @allele1 = split/��/, $allele{$num1[$z]}{$markers[$k]};
                        my @allele2 = split/��/, $allele{$num2[$z]}{$markers[$k]};
                        my @allele3 = split/��/, $allele{$num3[$z]}{$markers[$k]};
                        my @area3   = split/��/, $area{$num3[$z]}{$markers[$k]};
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

        ##############################################################################
                $worksheet->hide_gridlines();
                $worksheet->keep_leading_zeros();
                $worksheet->set_column(0,0,0.5);
                $worksheet->set_column(1,1,14.5);
                $worksheet->set_column(2,2,10);
                $worksheet->set_column(3,4,8.5);
                $worksheet->set_column(5,5,11);
                $worksheet->set_column(6,7,10.5);
                $worksheet->set_column(8,8,14);
                my @rows = (73,8,3,18.4,22.8,22.8,10,16.2,16.2,16.2,16.2,16.2,10,18.6,18.6,18.6,18.6,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,16.2,39,57.6,12.6,25,25);
                for my $i (0 .. $#rows){$worksheet->set_row($i, $rows[$i]);}

                $worksheet->set_margin_left(0.394);
                $worksheet->set_margin_right(0.394);
                $worksheet->set_margin_top(0.2);
                #my $footer = '&R'.decode('GB2312',$name{$num3[$z]}.'-'.$hospital{$num3[$z]}.'��'.$doctor{$num3[$z]}.'������').'&P'.decode('GB2312','ҳ/��').'&N'.decode('GB2312','ҳ');
                my $footer = '&L'.decode('GB2312','����ʵ���ң�ݶ˶���ϼ���ʵ����')."\n".
                             #'&R'.decode('GB2312','CSTB-B-R-0021-1.0')."\n".
                             '&L'.decode('GB2312','��ѯ�绰��0755-89323317').
                             '&R'.decode('GB2312',$name{$num3[$z]}.'����').'&P'.decode('GB2312','ҳ/��').'&N'.decode('GB2312','ҳ');

                $worksheet->set_footer($footer);
                $worksheet->insert_image('B1', "pic/ݶ˶����logo.png", 10, 10, 0.73, 0.73);

                $worksheet->merge_range('B1:I1', decode('GB2312','Ƕ��״̬��������'), $format1);
                $worksheet->merge_range('B3:I3', decode('GB2312',''), $format2);
                $worksheet->merge_range('G4:I4', decode('GB2312','���浥�ţ�'.$rptnum{$num3[$z]}.'   '),$format3);
                $worksheet->write('B05',decode('GB2312','�ͼ쵥λ'),$format4);
                $worksheet->write('B06',decode('GB2312','�����Ŀ'),$format4);
                $worksheet->merge_range('C5:G5', decode('GB2312',$hospital{$num3[$z]}),$format5);
                $worksheet->write('H05',decode('GB2312','�ͼ�ҽ��'),$format4);
                if ($doctor{$num3[$z]} =~ /^([\x80-\xFF][\x80-\xFF])$/){   #���� ����
                     $doctor{$num3[$z]} = $1.'ҽ��';
                     }
                $worksheet->write('I05',decode('GB2312',$doctor{$num3[$z]}),$format4);
                my $testt;
                if ($cells{$num3[$z]} =~ /(\S+)��ѡ/){
                    $testt = $1.'Ƕ��״̬����';
                    }
                else{
                    $testt = 'ȫѪǶ��״̬����';
                    }
                $worksheet->merge_range('C6:I6', decode('GB2312',$testt),$format5);
                $worksheet->merge_range('B8:I8', decode('GB2312','������Ϣ'),$format7);
                $worksheet->write('B09',decode('GB2312','�������'),$format6);
                $worksheet->write('C09',decode('GB2312','����'),$format6);
                $worksheet->write('D09',decode('GB2312','�Ա�'),$format6);
                $worksheet->write('E09',decode('GB2312','����'),$format6);
                $worksheet->write('F09',decode('GB2312','��������'),$format6);
                $worksheet->write('G09',decode('GB2312','��������'),$format6);
                $worksheet->write('H09',decode('GB2312','��������'),$format6);
                $worksheet->write('I09',decode('GB2312','��ϵ'),$format6);
                $worksheet->write('B10',decode('GB2312',$number{$num3[$z]}),$format10);
                $worksheet->write('B11',decode('GB2312',$number{$num2[$z]}),$format10);
                $worksheet->write('C10',decode('GB2312',$name{$num3[$z]}),$format6);
                $worksheet->write('C11',decode('GB2312',$name{$num2[$z]}),$format6);
                $worksheet->write('D10',decode('GB2312',$gender{$num3[$z]}),$format6);
                $worksheet->write('D11',decode('GB2312',$gender{$num2[$z]}),$format6);
                $worksheet->write('E10',decode('GB2312',$age{$num3[$z]}),$format6);
                $worksheet->write('E11',decode('GB2312',$age{$num2[$z]}),$format6);
                $worksheet->write('F10',decode('GB2312',$sample{$num3[$z]}),$format6);  #####################
                $worksheet->write('F11',decode('GB2312',$sample{$num2[$z]}),$format6);  #####################
                $worksheet->write('G10',decode('GB2312',DateUnify($date3{$num3[$z]})),$format10);
                $worksheet->write('G11',decode('GB2312',DateUnify($date3{$num2[$z]})),$format10);
                $worksheet->write('H10',decode('GB2312',DateUnify($date4{$num3[$z]})),$format10);
                $worksheet->write('H11',decode('GB2312',DateUnify($date4{$num2[$z]})),$format10);
                $worksheet->write('B12',decode('GB2312','סԺ/�����'),$format6);
                $worksheet->write('E12',decode('GB2312','����'),$format6);
                $worksheet->write('G12',decode('GB2312','�ٴ����'),$format6);
                $worksheet->merge_range('C12:D12', decode('GB2312',$hosptl_num{$num3[$z]}), $format7);
                $worksheet->write('F12',decode('GB2312',$bed_num{$num3[$z]}),$format6);

                if ($diagnosis{$num3[$z]} ne "-"){
                     $worksheet->merge_range('H12:I12', decode('GB2312',$diagnosis{$num3[$z]}), $format7);
                }else{
                     $worksheet->merge_range('H12:I12', decode('GB2312',$diagnosis{$num1[$z]}), $format7);
                     }

                my $tmp = $sheet[$z];
                if ($relation{$num2[$z]} =~ /$tmp/){
                        $relation{$num2[$z]} =~ s/$tmp//;
                }
                $worksheet->write('I10',decode('GB2312',$relation{$num3[$z]}),$format6);
                $worksheet->write('I11',decode('GB2312',$relation{$num2[$z]}),$format6);
                $worksheet->merge_range('B14:I14', decode('GB2312','�����'), $format5);
                $worksheet->merge_range('B15:B17', decode('GB2312','STRλ��'), $format15);
                $worksheet->merge_range('C15:H15', decode('GB2312','��λ����'), $format5);
                if ($sample{$num1[$z]} =~ /��ǻ/){
                        $worksheet->merge_range('C16:D16', decode('GB2312','������ֲǰ(��ǻ)'), $format7);
                }else{
                        $worksheet->merge_range('C16:D16', decode('GB2312','������ֲǰ'), $format7);
                }
                $worksheet->merge_range('E16:F16', decode('GB2312','��    ��'), $format7);
                $worksheet->merge_range('G16:H16', decode('GB2312','������ֲ��'), $format7);
                $worksheet->merge_range('C17:D17', decode('GB2312','������ţ�'.$num1[$z]), $format16);
                $worksheet->merge_range('E17:F17', decode('GB2312','������ţ�'.$num2[$z]), $format16);
                $worksheet->merge_range('G17:H17', decode('GB2312','������ţ�'.$num3[$z]), $format16);
                $worksheet->merge_range('I15:I17', decode('GB2312','λ��״̬'), $format15);
                for my $q (0..$#markers){
                        $worksheet->write($q+17,1,$markers[$q], $format6);
                        $worksheet->merge_range($q+17,2,$q+17,3,decode('GB2312',$allele{$num1[$z]}{$markers[$q]}), $format7);
                        $worksheet->merge_range($q+17,4,$q+17,5,decode('GB2312',$allele{$num2[$z]}{$markers[$q]}), $format7);
                        $worksheet->merge_range($q+17,6,$q+17,7,decode('GB2312',$allele{$num3[$z]}{$markers[$q]}), $format7);
                        $worksheet->write($q+17,8,decode('GB2312',$marker_type[$z][$q]), $format6);
                }
                $worksheet->write('B34',decode('GB2312','�����ۣ�'),$format17);
                if ($count_avg[$z] =~ /\d/){
                        $count_avg[$z] = sprintf("%.2f", $count_avg[$z]*100);
                        unless ($conclusion[$z]){
                                if($count_avg[$z] >= 95){
                                        $conclusion[$z] = '������ֲ����ϸ��ռ'.$count_avg[$z].'%������Ϊ��ȫǶ��״̬��';
                                        $worksheet->merge_range('C34:I34',decode('GB2312',$conclusion[$z]), $format18);
                                }elsif($count_avg[$z] < 5){
                                        $conclusion[$z] = '������ֲ����ϸ��ռ'.$count_avg[$z].'%������Ϊ΢Ƕ��״̬��';
                                        $worksheet->merge_range('C34:I34',decode('GB2312',$conclusion[$z]), $format18);
                                }else{
                                        $conclusion[$z] = '������ֲ����ϸ��ռ'.$count_avg[$z].'%������Ϊ���Ƕ��״̬��';
                                        $worksheet->merge_range('C34:I34',decode('GB2312',$conclusion[$z]), $format18);
                                }
                        }else{
                                $worksheet->merge_range('C34:I34',decode('GB2312',$conclusion[$z]), $format18);
                        }
                }else{
                        $worksheet->merge_range('C34:I34',decode('GB2312','��'), $format18);
                }
                $worksheet->write('B35',decode('GB2312','��    ע'),$format15);
                $worksheet->merge_range('C35:I35', decode('GB2312','1��Ƕ��״̬�綨[1]
��ȫǶ��״̬��CC��: DC��95%; ���Ƕ��״̬��MC��:5%��DC<95%�� ΢Ƕ��״̬��DC��5%��
[1] Outcome of patients with hemoglobinopathies given either cord blood or bone marrow
transplantation from an HLA-idebtucak sibling.Blood.2013,122(6):1072-1078.
2����������������ѧ���ݱȶԡ����������ٴ���ⱨ�档'), $format19);

                $worksheet->merge_range('B37:C37', decode('GB2312','��  ��  ��'), $format7);
                $worksheet->merge_range('B38:C38', decode('GB2312','��  ��  ��'), $format7);
                $worksheet->merge_range('D37:E37', decode('GB2312',''), $format7);
                $worksheet->merge_range('D38:E38', decode('GB2312',''), $format7);
                $worksheet->merge_range('F37:G37', decode('GB2312','�������'), $format7);
                $worksheet->merge_range('F38:G38', decode('GB2312','��������'), $format7);
                $worksheet->merge_range('H37:I37', decode('GB2312',DateUnify($date1{$num3[$z]})), $format9);
                $worksheet->merge_range('H38:I38', decode('GB2312',sprintf("%d-%02d-%02d",$year,$mon,$mday)), $format9);

                if (-e "pic/�����.png"){
                     $worksheet->insert_image('D37', "pic/�����.png", 5, 0, 1, 1);
                     }
                if (-e "pic/������.png"){
                     $worksheet->insert_image('D38', "pic/������.png", 20, 0, 1, 1);
                     }
                if (-e "pic/����.png"){
                     $worksheet->insert_image('H37', "pic/����.png", 10, 12, 1, 1);
                     }


#����        ҽԺ        ��������        �������        ������        Ƕ����

                if ($count_avg[$z] == 0){
                        printf SUM "%s\t%s\t%s\t%s\t%s\t%f%s\t%d\tNA\tNA\n", $name{$num3[$z]}, $hospital{$num3[$z]}, $sample{$num3[$z]}, $number{$num3[$z]}, $rptnum{$num3[$z]}, $count_avg[$z],"%",$count_n[$z];
                }
                else{
#����\tҽԺ\t��������\t�������\t������\tǶ����\t��Чλ��\tSD\tCV
                        printf SUM "%s\t%s\t%s\t%s\t%s\t%f%s\t%d\t%.2f%s\t%.2f%s\n", $name{$num3[$z]}, $hospital{$num3[$z]}, $sample{$num3[$z]}, $number{$num3[$z]}, $rptnum{$num3[$z]}, $count_avg[$z],"%",$count_n[$z], $SD[$z]*100,"%", $SD[$z]/$count_avg[$z]*10000,"%";
                }

#
                if ($exp_num{$TCAID} != 3 or $sheet_name{$sheet[$z]} > 1){
                        $workbook -> close();
                        $RptBox -> Append("�������ɳɹ�����Ƕ������\r\n");
                        next;
                }
                $RptBox -> Append("�������ɳɹ���");

        #####################################################################################
                my $tempid = $identity{$TCAID};
                my $i;
                my $j = 1;
                my $Chart_Marker_Num = 0;
                my %Graphic_SampleID;
                my %Graphic_Chimerism;
                my %Types;
                my @date_seq;
                push @date_seq, 0;
                if(exists $Chimerism{$tempid}){
                        foreach $i(0 .. $#{$Chimerism{$tempid}}){
                                my $Chmrsm = $Chimerism{$tempid}[$i];
                                $Chmrsm =~ s/%//;
                                $Chmrsm = sprintf ("%.2f", $Chmrsm);
                                my $Smplid = $SampleID{$tempid}[$i];
                                my $SmpType = $sampleType{$Smplid};
                                next unless $SmpType;
                                next if $SmpType eq "-";
                                my $rptDate = DateUnify($ReportDate{$tempid}[$i]);
                                my $rcvDate = DateUnify($receiveDate{$Smplid});
                                my $smplDate = DateUnify($sampleDate{$Smplid});
                                my $tmpDate;
                                if ($smplDate ne '����' && $smplDate ne '-'){
                                        $tmpDate = $smplDate;
                                }elsif ($rcvDate ne '����' && $rcvDate ne '-'){
                                        $tmpDate = $rcvDate;
                                }elsif($rptDate ne '����'){
                                        $tmpDate = $rptDate;
                                }else{
                                        $tmpDate = sprintf "%s%d%s", "����", $j, "��";
                                }

                                # print $Smplid,"|", $rptDate,"|",$rcvDate,"|",$smplDate,"|",$tmpDate,"\n";

                                $Graphic_Chimerism{$tmpDate}{$SmpType} = $Chmrsm;
                                $Graphic_SampleID{$tmpDate}{$SmpType} = $Smplid;
                                $Types{$SmpType} ++;
                                if ($date_seq[-1] ne $tmpDate || $tmpDate =~ /����/){
                                        push @date_seq, $tmpDate;
                                        $j ++;
                                }
                                $Chart_Marker_Num ++;
                        }
                        shift @date_seq;
                        my $headings;
                        push @{$headings}, decode('GB2312', 'ʱ��');
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

                        $graphic_temp -> write('A1', $headings);
                        $graphic_temp -> write('A2', $write_data);
                        $graphic->hide_gridlines();
                        $graphic->keep_leading_zeros();
                        $graphic->set_column(0,0,4.24);
                        $graphic->set_column(1,6,13.75);
                        $graphic->set_column(7,7,4.24);
                        my @rows = (75,   3.6, 3.6 , 3.6 , 19 ,  18 ,  15.6, 15.6,
                                                15.6, 18.6,        25.8, 18.6, 19.2, 24.6, 19.8, 16.2,
                                                16.2, 16.2, 16.2, 16.2,        16.2, 16.2, 16.2, 24, #24 from 32 to 24
                                                18.75,16.25,16.25,16.25,33.0, 16.25, 16.25);#last from 16.25 to 17
                        for my $i(0 .. $#rows){
                                $graphic->set_row($i, $rows[$i]);
                        }
                        foreach $i(1 .. $Chart_Marker_Num){
                                $graphic->set_row($i+30, 13.5);
                        }
                        $graphic->set_margin_left(0.394);
                        $graphic->set_margin_right(0.394);
                        #$graphic->set_margin_top(0.2);
                        $graphic->set_footer($footer);
                        $graphic->insert_image('B1', "pic/ݶ˶����logo.png", 10, 10, 0.73, 0.73);

                        $graphic->merge_range('B1:G1', decode('GB2312','Ƕ������'), $format1);
                        #$graphic->merge_range('B2:G2', decode('GB2312','��ַ���Ϻ����ֶ�������Ƽ·908Ū21�ţ��Ϻ�����ҽѧ԰����          �ʱࣺ201318'), $format2);
                        $graphic->merge_range('B4:G4', decode('GB2312',''), $format2);
                        $graphic->write('B5',decode('GB2312','��������'), $Gfmt1);
                        $graphic->write('C5',decode('GB2312',$name{$num3[$z]}), $Gfmt2);
                        $graphic->write('F5',decode('GB2312','������ţ�'), $Gfmt1);
                        $graphic->write('G5',decode('GB2312',$number{$num3[$z]}), $Gfmt1);
                        my $chart = $workbook->add_chart(type => 'line', embedded => 1 );

                        my $row_max = $#{${$write_data}[0]}+1;
                        my $col_max = $#{$write_data};


                        for my $i(1..$col_max){
                                my $formula = sprintf "=temp!\$%s1", chr($i+65);
                                $chart->add_series(
                                        categories => ['temp', 1,$row_max, 0 , 0],
                                        values     => ['temp', 1, $row_max, $i, $i],
                                        name_formula => $formula,
                                        # name       => decode('GB2312',${$headings}[$i]),
                                                marker   => {
                                                        type    => 'automatic',
                                                        size    => 5,
                                                },
                                );
                        }

                        $chart->set_chartarea(
                                color => 'white',
                                line_color => 'black',
                                line_weight => 3,
                        );

                        $chart->set_plotarea(
                                color => 'white',

                        );

                        $chart->set_y_axis(
                                name => decode('GB2312','Ƕ����(%)'),
                                min  => 0,
                                max  => 100,
                                major_unit => 20,
                        );


                        $chart->set_legend( position => 'bottom' );
                        $chart->set_size( width => 607, height => 400 );
                        $graphic->insert_chart('B7', $chart);

                        $graphic->merge_range('B25:G25', decode('GB2312','TCA���ڼ������'), $Gfmt5);
                        $graphic->merge_range('B26:G26', decode('GB2312','���߼�⣺��ǰͬʱ�Թ������߽��м�⡢Ҳ�����������״�׷�ټ��ʱ����'), $Gfmt6);
                        $graphic->merge_range('B27:G27', decode('GB2312','׷�ټ�⣺����������2�ܽ��е�һ��TCA����4�ܽ��еڶ��μ�⣻'),$Gfmt6);
                        $graphic->merge_range('B28:G28', decode('GB2312','        ����6�����ڣ�ÿ�¼��һ�Σ�6����֮��ÿ2���¼��һ�Σ�ֱ��Ƕ�����ȶ�'), $Gfmt6);
                        $graphic->insert_image('B29', "pic/comment.bmp", 5, 5);
                        $graphic->merge_range('B30:G30', decode('GB2312','��ܰ��ʾ��һ�������������Ʒ����������ڵ�����2����Ҫ�����������'), $Gfmt7);

                        $graphic->write('B32', decode('GB2312','������'), $Gfmt3);
                        $graphic->write('C32', decode('GB2312','��������'), $Gfmt3);
                        $graphic->write('D32', decode('GB2312','�������'), $Gfmt3);
                        $graphic->write('E32', decode('GB2312','Ƕ����(%)'),   $Gfmt3);
                        $graphic->write('F32', decode('GB2312','�������'), $Gfmt3);
                        $graphic->write('G32', decode('GB2312','��������'), $Gfmt3);

                        my $i = 1;
                        my $j = 1;
                        for my $tmpDate(@date_seq){
                                for my $SmpType(keys %Types){
                                        my $Smplid = $Graphic_SampleID{$tmpDate}{$SmpType};
                                        my $Chmrsm = $Graphic_Chimerism{$tmpDate}{$SmpType};
                                        next unless $Smplid;
                                        my $rcvDate = $receiveDate{$Smplid};
                                        my $smplDate = $sampleDate{$Smplid};
                                        $graphic->write($j+31, 1, $i, $Gfmt4);
                                        $graphic->write($j+31, 2, decode('GB2312',$smplDate), $Gfmt4);
                                        $graphic->write($j+31, 3, decode('GB2312',$rcvDate), $Gfmt4);
                                        $graphic->write($j+31, 4, sprintf("%.2f",$Chmrsm), $Gfmt8);
                                        $graphic->write($j+31, 5, $Smplid, $Gfmt4);
                                        $graphic->write($j+31, 6, decode('GB2312',$SmpType), $Gfmt9);
                                        $j ++;
                                        if (($j-11)%54 == 0){
                                                $graphic->write($j+31, 1, decode('GB2312','������'), $Gfmt3);
                                                $graphic->write($j+31, 2, decode('GB2312','��������'), $Gfmt3);
                                                $graphic->write($j+31, 3, decode('GB2312','�������'), $Gfmt3);
                                                $graphic->write($j+31, 4, decode('GB2312','Ƕ����(%)'),   $Gfmt3);
                                                $graphic->write($j+31, 5, decode('GB2312','�������'), $Gfmt3);
                                                $graphic->write($j+31, 6, decode('GB2312','��������'), $Gfmt3);
                                                $j ++;
                                        }
                                }
                                $i ++;
                        }

                }else{

                }

                $workbook->close();
                $RptBox -> Append("Ƕ���������ɳɹ���\r\n");

        }


        $sb->Move( 0, ($main->ScaleHeight() - $sb->Height()) );
        $sb->Resize( $main->ScaleWidth(), $sb->Height() );
        if ($success){
                $sb->Text("������");
        }else{
                $sb->Text("�����ɣ��д���");
        }

        $RUNwindow -> Hide();
        if ($success){
                $error =  "�������ɹ���\n";
                Win32::MsgBox $error, 0, "�ɹ���";
        }else{
                $error =  "�������ɹ����������˴���\n";
                Win32::MsgBox $error, 0, "ע�⣡";
        }

        close SUM;

}

sub RUN_MouseMove{
        $sb -> Text('���У�������������');
}

sub RUN_MouseOut{
        $sb -> Text('');
}

sub QUIT_MouseMove{
        $sb -> Text('�˳�');
}

sub QUIT_MouseOut{
        $sb -> Text('');
}

sub QUIT_Click{
        &WriteConfig;

        return -1;
}

sub COPY_MouseMove{
        $sb -> Text('�����Ҳ�������м�¼��������');
}

sub COPY_MouseOut{
        $sb -> Text('');
}

sub COPY_Click{
        $RptBox -> SelectAll();
        $RptBox -> Copy();
        $error = '�ɹ������������壡';
        Win32::MsgBox $error, 0 ,"�ɹ���";
}

sub Shorten{
        my ($string, $lim) = @_;
        if ($string =~ /$pwd\\(.+)$/){
                $string = ".\\".$1;
        }
        my $len = length($string);
        return $string if $len <= $lim;
        $string =~ /^.+\\([^\\]+)$/;
        $len = length($1);
        my $tmp = sprintf "%s...\\%s", substr($string, 0 ,$lim-$len-4), $1;
        return $tmp;
}

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

sub WriteConfig{
        `attrib -r -h TCAconfig.ini` if (-e "TCAconfig.ini");
        open IN,"> TCAconfig.ini";
        foreach (@ConfigList){
                print IN $_,"\t",$ConfigHash{$_},"\n";
        }
        close IN;
        `attrib +r +h TCAconfig.ini`;
}

sub DateUnify{
        return $_[0] if $_[0] eq '����';

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

sub Avg_SD{
        my $total = 0;
        my $SD = 0;

        $total += $_ foreach @_;
        $total /= @_;

        $SD += ($total-$_)*($total-$_) foreach @_;
        $SD = sqrt($SD/@_);

        return ($total, $SD);
}