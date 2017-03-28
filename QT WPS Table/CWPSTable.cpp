#include "CWPSTable.h"
#include <QDebug>
# pragma execution_character_set("utf-8")

CWPSTable::CWPSTable(bool visible,bool alert) {//???
	if(!newExcel()) return;
	setAlert(alert);
	setVisible(visible);
	//MyExcel.setCaption("t");// TITLE
	getAddWorkBook();
	getSheets();
	getSheet(1);
	//setScreenUpdating(true);
}

CWPSTable::CWPSTable(const QString& openPath,bool visible,bool alert) {
	if(!newExcel()) return;
	setAlert(alert);
	setVisible(visible);
	getOpenWorkBook(openPath);
	getSheets();
	getSheet(1);
	//setScreenUpdating(true);
}

CWPSTable::~CWPSTable() {
	qDebug() << "release MyWPS";
}

void CWPSTable::getAddWorkBook() {
	getWorkBooks();
	addWorkBooks();//if create NOT open, got to add a new workBook
	activateCurrWorkBook();
}

void CWPSTable::getOpenWorkBook(const QString& openPath) {
	getWorkBooks();
	openWorkBooks(openPath);
	activateCurrWorkBook();
}

bool CWPSTable::newExcel() {
	MyExcel = new QAxObject("Ket.Application");
	if(MyExcel->isNull()) {
		return false;
	}
	return true;
}

void CWPSTable::deleteExcel() {
	delete MyExcel;
	MyExcel = nullptr;
}

void CWPSTable::save_Quit(const QString& path) {
	saveAs(path);
	quit();
	deleteExcel();
}

void CWPSTable::openWorkBooks(QString inOpenPath) {
	//	workbooks->querySubObject("Open(QString,QVariant,QVariant)", fileName, 3, true);
	//	MyWorkBook = MyWorkBooks->querySubObject("Open(QString,QVariant,QVariant)", inOpenPath, 3, true);
	//	MyWorkBook->querySubObject("Open (const QString&)", inOpenPath);
	//
	//example:dynamicCall("Open (const QString&)", "C:/test.xls");
	MyWorkBooks->dynamicCall("Open (const QString&)", inOpenPath);
	if(MyWorkBooks->isNull()) {
		qDebug() << "指定的excel文件不存在！";
	}
}

void CWPSTable::setVisible(bool inBool) {
	MyExcel->setProperty("Visible", inBool);
}

void CWPSTable::setAlert(bool inBool) {
	MyExcel->setProperty("DisplayAlerts", inBool);
}

void CWPSTable::setCaption(QString inCaption) {
	MyExcel->setProperty("Caption", inCaption);
}

void CWPSTable::getWorkBooks() {
	MyWorkBooks = MyExcel->querySubObject("Workbooks");
	if(MyWorkBooks->isNull()) {//???
		qDebug() << "No WorkBooks found~";
	}
}

int CWPSTable::countSheets() {
	return mySheets->property("Count").toInt();//获取表数量
}

void CWPSTable::getSheets() {
	//mySheets = MyWorkBook->querySubObject("Sheets"); workable
	mySheets = MyWorkBook->querySubObject("Worksheets");
}

void CWPSTable::addSheet(bool addAtLast) {//the sheet is got after adding
	if(!addAtLast) {
		mySheets->dynamicCall("Add");
	} else {
		int SheetsCnt = countSheets();//获取表数量
		QAxObject* pLastSheet = mySheets->querySubObject("Item(int)", SheetsCnt);//这步可以的
		mySheets->querySubObject("Add(QVariant)", pLastSheet->asVariant());//这步之后，新加表 处于倒数第二的位置；
		mySheet = mySheets->querySubObject("Item(int)", SheetsCnt);
		pLastSheet->dynamicCall("Move(QVariant)", mySheet->asVariant());
	}
}
/**
 * \brief  ActiveWindow.SelectedSheets.Copy After:=Sheets("Sheet1")
 * \param addAtLast 
 */
void CWPSTable::copySheet() {
//	QAxObject* lastSheet = mySheets->querySubObject("Item(int)", countSheets());//workable
//	mySheet->dynamicCall("Copy(const QVariant&)", lastSheet->asVariant());//new sheet at last 2 pos
	mySheet->dynamicCall("Copy(const QVariant&)",mySheet->asVariant());//add sheet before itself,but how to add after?
}
void CWPSTable::CopyWorkBook() {//copyALlSheets
	mySheets->dynamicCall("Copy()");
}


void CWPSTable::addWorkBooks() {
	MyWorkBooks->dynamicCall("Add");
}

void CWPSTable::activateCurrWorkBook() {
	MyWorkBook = MyExcel->querySubObject("ActiveWorkBook");
}

void CWPSTable::getSheet(int inIndex) {
	//mySheet = MyWorkBook->querySubObject("Worksheets(int)", inIndex);//workable
	//	mySheet = MyWorkBook->querySubObject("Worksheets(const QVariant&)", "酒店");//workable
	mySheet = MyWorkBook->querySubObject("Sheets(int)", inIndex);
	if(mySheet->isNull()) {
		qDebug() << "No Sheet found~";
	}
	mySheet->dynamicCall("activate()");
}
void CWPSTable::setSheetName(const QString& name) {
	 mySheet->setProperty("Name",name);
}
QString CWPSTable::getSheetName() {
	return mySheet->property("Name").toString();
}

void CWPSTable::getRange(QString inRange) {
	range = mySheet->querySubObject("Range(const QVariant&)", QVariant(inRange));
}

void CWPSTable::getRange(int inRow1,int inColumn1) {
	QAxObject* myCell1 = mySheet->querySubObject("Cells(int,int)", inRow1, inColumn1);
	range = mySheet->querySubObject
			("Range(const QVariant&,const QVariant&)", myCell1->asVariant(), myCell1->asVariant());
	range->dynamicCall("Select()");
}

QAxObject* CWPSTable::getRange(int inRow1,int inColumn1,int inRow2,int inColumn2) {
	QAxObject* myCell1 = mySheet->querySubObject("Cells(int,int)", inRow1, inColumn1);
	QAxObject* myCell2 = mySheet->querySubObject("Cells(int,int)", inRow2, inColumn2);
	return range = mySheet->querySubObject
			("Range(const QVariant&,const QVariant&)", myCell1->asVariant(), myCell2->asVariant());
}

QString CWPSTable::getValue(int inRow1,int inColumn1) {
	getRange(inRow1, inColumn1);
	return range->property("Value2").toString();
}

//void CWPSTable::getRange(int inRow1, int inColumn1, int inRow2, int inColumn2)
//{
//}
//
//void CWPSTable::getCell(int inRow, int inColumn)
//{
//	MyCell = mySheet->querySubObject("Cells(int,int)", inRow, inColumn);
//}

QString CWPSTable::getValue() {
	return range->property("Value2").toString();
}

void CWPSTable::setValue(const QVariant& inValue) {
	//	inRange->dynamicCall("SetValue(const QVariant&)", QObject::tr("卡卡尼莫"));
	//	getRange(inRange);
	range->dynamicCall("SetValue(const QVariant&)", inValue);
}

void CWPSTable::setValue(int row,int col,const QVariant& value) {
	getRange(row, col);
	setValue(value);
}

void CWPSTable::setComment(int row,int col,const QString& text) {
	getRange(row, col);
	range->dynamicCall("AddComment(const QVariant&)", text);
}

void CWPSTable::setRangeFormat(int row1,int col1,int row2,int col2,QString myFormat) {
	//"yyy年mm月dd";"@" '为文本 "G/通用格式" '为通用
	getRange(row1, col1, row2, col2);
	range->setProperty("NumberFormat", myFormat);
}

void CWPSTable::autoFit() {
	//selectRange.EntireColumn.AutoFit(); //全部列自适应宽度
	//selectRange.EntireRow.AutoFit();    //全部行自适应高度
	autoFitColumn();
	autoFitRow();
}

void CWPSTable::autoFitColumn() {
	mySheet->querySubObject("columns")->dynamicCall("AutoFit");//workable
}

void CWPSTable::autoFitRow() {
	mySheet->querySubObject("rows")->dynamicCall("AutoFit");//workable
}

//QString CWPSTable::getCellValue(int inRow, int inColumn)
//{
//	getCell(inRow, inColumn);
//	return	MyCell->property("Value2").toString();
//}
//
//void CWPSTable::setCellValue(int inRow, int inColumn, QString inValue)
//{
//	getCell(inRow, inColumn);
//	MyCell->querySubObject("SetValue(const QVariant&)", inValue);
//}
void CWPSTable::setBackColor(int inRow1,int inColumn1,int inRow2,int inColumn2,int inFontColor) {
	getRange(inRow1, inColumn1, inRow2, inColumn2);
	range->querySubObject("interior")->setProperty("ColorIndex", inFontColor);
}

void CWPSTable::setFontColor(int inRow,int inColumn,int inFontColor) {
	getRange(inRow, inColumn);
	range->querySubObject("Font")->setProperty("ColorIndex", inFontColor);
	//	MyRange->querySubObject("Font")->setProperty("Color", QColor(255, 0, 0));
}

void CWPSTable::setFontColor(int inRow1,int inColumn1,int inRow2,int inColumn2,int inFontColor) {
	getRange(inRow1, inColumn1, inRow2, inColumn2);
	range->querySubObject("Font")->setProperty("ColorIndex", inFontColor);
}

void CWPSTable::copyPaste(int inRow1,int inColumn1,int inRow2,int inColumn2,
                       int inPasteRow1,int inPasteColumn1,int inPasteRow2,int inPasteColumn2) {
	QAxObject* copyRange = getRange(inRow1, inColumn1, inRow2, inColumn2);
	getRange(inPasteRow1, inPasteColumn1, inPasteRow2, inPasteColumn2);
	copyRange->querySubObject("Copy(const QVariant&)", range->asVariant());
}

void CWPSTable::copyPasteFromSheet(int copySheetPos,int pasteSheetPos,
                                int inRow1,int inColumn1,int inRow2,int inColumn2,
                                int inPasteRow1,int inPasteColumn1) {
	getSheet(copySheetPos);
	QAxObject* copyRange = getRange(inRow1, inColumn1, inRow2, inColumn2);
	getSheet(pasteSheetPos);
	getRange(inPasteRow1, inPasteColumn1, inPasteRow1, inPasteColumn1);
	copyRange->querySubObject("Copy(const QVariant&)", range->asVariant());
}



/**
 * \brief 
 ? Is there a faster way to do it?
 */
void CWPSTable::copyInsertRowTo(int inRow1,int inColumn1,int inRow2,int inColumn2,
                             int inPasteRow,int inPasteColumn,bool inIsDown,int inPasteRowNum) {
	//	QAxObject* myCopyRange = getRange(inRow1, inColumn1, inRow2, inColumn2);
	//	myCopyRange->querySubObject("Select()");
	//	myCopyRange->querySubObject("Copy(const QVariant&)", NULL);
	getRange(inRow1, inColumn1, inRow2, inColumn2);
	//	MyRange->querySubObject("Select()");
	range->querySubObject("Copy(const QVariant&)", NULL);
	insertRow(inPasteRow, inPasteColumn, inIsDown, inPasteRowNum);
}

void CWPSTable::insertRow(int inRow,int inColumn,bool inIsDown,int inRowNum) {
	if(inIsDown) inRow++;
	getRange(inRow, inColumn, inRow + inRowNum - 1, inColumn);
	range->querySubObject("EntireRow")->querySubObject("Insert()");
	//	getRange(inRow, inColumn);
	//
	//	for (size_t i = 0; i < inRowNum; ++i)
	//	{
	//		MyRange->querySubObject("EntireRow")->querySubObject("Insert()");
	//	}
}

void CWPSTable::setScreenUpdating(bool inBool) {
	MyExcel->dynamicCall("SetScreenUpdating(bool)", inBool);
}

void CWPSTable::saveAs(QString inSavePath) {
	MyWorkBook->querySubObject("SaveAs(const QString&)", inSavePath);
}

void CWPSTable::quit() {
	MyWorkBook->dynamicCall("Close (Boolean)", false);
	MyExcel->dynamicCall("Quit()");
}

void CWPSTable::setColumnWidth(int col,float width) {
	mySheet->querySubObject("columns(int)", col)->setProperty("ColumnWidth", width);//workable
}

void CWPSTable::setColumnHidden(int col,bool isHidden) {
	mySheet->querySubObject("columns(int)", col)->setProperty("Hidden", isHidden);//workable
}

void CWPSTable::mergeRange(int row1,int col1,int row2,int col2) {
	this->getRange(row1, col1, row2, col2);
	range->querySubObject("Merge");
}

void CWPSTable::setRangeFontSize(int row1,int col1,int row2,int col2,int myFontSize) {
	this->getRange(row1, col1, row2, col2);
	range->querySubObject("Font")->setProperty("Size", myFontSize);
}

void CWPSTable::setRangeAlign(int row1,int col1,int row2,int col2,int alignType) {
	this->getRange(row1, col1, row2, col2);//.HorizontalAlignment = xlCenter
	//.VerticalAlignment = xlCenter=-4108
	range->setProperty("VerticalAlignment", alignType);
	range->setProperty("HorizontalAlignment", alignType);
}

void CWPSTable::setRangeFontStyle(int row1,int col1,int row2,int col2,int myStyle,bool styleFlag) {
	this->getRange(row1, col1, row2, col2);//.HorizontalAlignment = xlCenter
	switch(myStyle) {
	case 0:
		range->querySubObject("Font")->setProperty("Bold", styleFlag);
		break;
	case 1:
		range->querySubObject("Font")->setProperty("Italic", styleFlag);
		break;
	case 2:
		range->querySubObject("Font")->setProperty("Underline", styleFlag);
		break;
	default:
		return;
	}
}


void CWPSTable::setRangeBorderStyle_Width(int row1, int col1, int row2, int col2,int borderStyle,int borderWidth ) {
	this->getRange(row1, col1, row2, col2);
	range->querySubObject("Borders")->setProperty("LineStyle", borderStyle);
	range->querySubObject("Borders")->setProperty("Weight", borderWidth);
}

void CWPSTable::multiSelectCopy() {//??? NOT WORKABLE
	//VBA:
	//Range("C4,E4,G4").Select
	//Range("G4").Activate
	//Selection.Copy
	//Range("C26").Select
	//ActiveSheet.Paste
	//------
	//Range("F1").Select
	//Range(Selection, Selection.End(xlDown)).Select
	//Selection.Copy
	//Range("I29").Select
	//ActiveSheet.Paste
	//==============
//	QAxObject* pasteRange = getRange(2, 14, 2, 14);
//	getRange(1, 1, 1, 1);
//	range->dynamicCall("Select");
//	range->dynamicCall("Copy");
//	getRange(1, 2, 1, 2);
//		range->dynamicCall("Select");
//	range->dynamicCall("Copy");
////	range->dynamicCall("Select");
////	range->dynamicCall("Paste");
//	range->querySubObject("Copy(const QVariant&)", pasteRange->asVariant());
	
	//getRange(2, 3);
	MyExcel->querySubObject("Range(QVariant&,QVariant&)","C4","D4")->dynamicCall("Select()");
	QAxObject* selObj = MyExcel->querySubObject("Selection");//ok
	QVariant selObj2 = selObj->dynamicCall("End(int)", -4121);
	MyExcel->querySubObject("Range(QVariant&,QVariant&)", selObj->asVariant(),
		selObj2)->dynamicCall("Select()");

	MyExcel->querySubObject("Selection")->dynamicCall("Copy()");
	getRange(30, 10);
	range->dynamicCall("Select()");
	MyExcel->querySubObject("ActiveSheet")->dynamicCall("Paste()");
}