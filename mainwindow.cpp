#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <ActiveQt/qaxobject.h>
#include <ActiveQt/qaxbase.h>
#include <QFileDialog>

MainWindow::MainWindow(QWidget *parent): QMainWindow(parent), ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    rows=0;
    columns=0;
    model = new QStandardItemModel;
    ui->tableView->setModel(model);

    connect(ui->applyButton, &QPushButton::clicked, this, &MainWindow::applyButtonClicked);
    connect(ui->exportButton, &QPushButton::clicked, this, &MainWindow::exportButtonClicked);
}

void MainWindow::applyButtonClicked()
{
    /**
    * @brief Обработчик нажания на кнопку создания таблицы
    */
    model = new QStandardItemModel;
    columns = ui->columnsLineEdit->text().toInt();
    rows = ui->rowsLineEdit->text().toInt();
    for(int i = 0; i < rows; i++)
    {
        QList<QStandardItem*> list;
        for(int j = 0; j < columns; j++)
            list.push_back(new QStandardItem);
        model->appendRow(list);
    }
    ui->tableView->setModel(model);
}

void MainWindow::exportButtonClicked()
{
    /**
    * @brief Обработчик нажания на иконку экспорта таблицы
    */
    if(rows<=0 || columns <=0)
        return;
    QString path = QFileDialog::getSaveFileName();
    if(!path.endsWith(".xlsx"))
        path += ".xlsx";
    QFile exportFile(path);
    if(!exportFile.exists())
    {
        QFile::copy(":/files/export.xlsx", path);
        exportFile.setPermissions(QFileDevice::ReadOwner | QFileDevice::WriteOwner);
        QFileDevice::Permissions p = exportFile.permissions();
        p.setFlag(QFileDevice::WriteOwner, true);
        p.setFlag(QFileDevice::WriteGroup, true);
        p.setFlag(QFileDevice::WriteOwner, true);
        p.setFlag(QFileDevice::WriteUser, true);
    }
    QString filename = path;
    filename = QString(filename.remove(0, filename.lastIndexOf('/')+1));

    QAxObject *mExcel = new QAxObject("Excel.Application", this);
    QAxObject *workbooks = mExcel->querySubObject("Workbooks");
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", path);
    QAxObject *mSheets = workbook->querySubObject("Sheets");
    QAxObject *StatSheet = mSheets->querySubObject("Item(const QVariant&)", QVariant("Sheet1"));

    QAxObject* Cell1 = StatSheet->querySubObject("Cells(QVariant&,QVariant&)", 1, 1);
    QAxObject* Cell2 = StatSheet->querySubObject("Cells(QVariant&,QVariant&)", rows, columns);
    QAxObject* range = StatSheet->querySubObject("Range(const QVariant&,const QVariant&)", Cell1->asVariant(), Cell2->asVariant() );
    QList<QVariant> cellsList;
    QList<QVariant> rowsList;
    for (int i = 0; i < rows; i++)
    {
        cellsList.clear();
        for(int j = 0; j < columns; j++)
            cellsList << ui->tableView->model()->data(ui->tableView->model()->index(i, j));
        rowsList << QVariant(cellsList);
    }
    range->setProperty("Value", QVariant(rowsList));
    delete range;
    delete Cell1;
    delete Cell2;

    QVariantList params = {QVariant(true), QVariant(filename), QVariant(false)};
    workbook->dynamicCall("Close(QVariant)", params);
    mExcel->dynamicCall( "Quit()");
}

MainWindow::~MainWindow()
{
    delete ui;
}

