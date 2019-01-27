#include "widget.h"
#include "ui_widget.h"
//2019-1-27
Widget::Widget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::Widget)
{
    ui->setupUi(this);
    //pApplication = NULL;
    //pWorkBooks = NULL;
    //pWorkBook = NULL;
    //pSheets = NULL;
    //pSheet = NULL;

    Today=QDate::currentDate();
    QString now = Today.toString("yyyy/MM/dd");
    ui->Date->setText(now);
    QSettings qSetting ("sys.ini",QSettings::IniFormat);
    ui->ID->setText(qSetting.value("sys/id").toString());
    ui->Name->setText(qSetting.value("sys/name").toString());
    ui->Level->setText(qSetting.value("sys/level").toString());
    ui->Department->setText(qSetting.value("sys/department").toString());
    ui->Office->setText(qSetting.value("sys/office").toString());
    ui->Post->setText(qSetting.value("sys/post").toString());
    ui->Btime->setText(qSetting.value("sys/btime").toString());
    ui->ETime->setText(qSetting.value("sys/etime").toString());
    ui->hours->setText(qSetting.value("sys/hours").toString());
    ui->Why->setText(qSetting.value("sys/why").toString());
    ui->Location->setText(qSetting.value("sys/loaction").toString());
}

Widget::~Widget()
{
    delete ui;
}

void Widget::on_Online_clicked()//提报按钮
{
    QSettings setting ("path.ini",QSettings::IniFormat);//设置路径文件
    QString path,lastpath;
    lastpath = setting.value("lastfilepath").toString();//保存路径
    path = QFileDialog::getOpenFileName(this,tr("Open File"),lastpath,tr("Model File(*.xl*)"));//获得上次路径
    setting.setValue("lastfilepath",path);
    QFileInfo fi;
    fi = QFileInfo(path);
    path = fi.filePath();
    ui->Path->setText(path);//读取路径

    QAxObject excel("Excel.Application");
    excel.setProperty("Visible",true);
    QAxObject *workbooks = excel.querySubObject("WorkBooks"); //获取工作簿集合
    workbooks->dynamicCall("Open(const QString&)",path);                //打开Excel文件
    QAxObject *workbook = excel.querySubObject("ActiveWorkBook");        //获取活动工作簿
    QAxObject *worksheets = workbook->querySubObject("Sheets");    //获取所有的工作表
    QAxObject *worksheet = worksheets->querySubObject("Item(int)",2);
    QAxObject *EDate;
    QAxObject *EID;
    QAxObject *EName;
    QAxObject *ELevel;
    QAxObject *EDepartment;
    QAxObject *EOffice;
    QAxObject *EPost;
    QAxObject *EBtime;
    QAxObject *EETime;
    QAxObject *Ehours;
    QAxObject *EWhy;
    QAxObject *ELocation;
    QString CD;
    int i =0;
    do
    {
        i++;
        EDate = worksheet->querySubObject("Cells(int,int)",i,1);
        CD = EDate->dynamicCall("Value2()").toString();
    }while(!(CD == NULL));
    EID = worksheet->querySubObject("Cells(int,int)",i,2);
    EName = worksheet->querySubObject("Cells(int,int)",i,3);
    ELevel = worksheet->querySubObject("Cells(int,int)",i,4);
    EDepartment = worksheet->querySubObject("Cells(int,int)",i,5);
    EOffice = worksheet->querySubObject("Cells(int,int)",i,6);
    EPost = worksheet->querySubObject("Cells(int,int)",i,7);
    EBtime = worksheet->querySubObject("Cells(int,int)",i,8);
    EETime = worksheet->querySubObject("Cells(int,int)",i,9);
    Ehours = worksheet->querySubObject("Cells(int,int)",i,10);
    EWhy = worksheet->querySubObject("Cells(int,int)",i,11);
    ELocation = worksheet->querySubObject("Cells(int,int)",i,12);

    EDate->setProperty("Value",ui->Date->text());
    EID->setProperty("Value",ui->ID->text());
    EName->setProperty("Value",ui->Name->text());
    ELevel->setProperty("Value",ui->Level->text());
    EDepartment->setProperty("Value",ui->Department->text());
    EOffice->setProperty("Value",ui->Office->text());
    EPost->setProperty("Value",ui->Post->text());
    EBtime->setProperty("Value",ui->Btime->text());
    EETime->setProperty("Value",ui->ETime->text());
    Ehours->setProperty("Value",ui->hours->text());
    EWhy->setProperty("Value",ui->Why->text());
    ELocation->setProperty("Value",ui->Location->text());
    workbook->dynamicCall("Save()");
    //excel.dynamicCall("Quit(void)");

}

void Widget::on_Opne_clicked()//保存按钮
{
    QSettings qSetting ("sys.ini",QSettings::IniFormat);
    qSetting.beginGroup("sys");
    qSetting.setValue("id",ui->ID->text());
    qSetting.setValue("name",ui->Name->text());
    qSetting.setValue("level",ui->Level->text());
    qSetting.setValue("department",ui->Department->text());
    qSetting.setValue("office",ui->Office->text());
    qSetting.setValue("post",ui->Post->text());
    qSetting.setValue("btime",ui->Btime->text());
    qSetting.setValue("etime",ui->ETime->text());
    qSetting.setValue("hours",ui->hours->text());
    qSetting.setValue("why",ui->Why->text());
    qSetting.setValue("loaction",ui->Location->text());
}
