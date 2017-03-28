#pragma once

#define RED 3
#define YELLOW 6
#define GREEN 43
#define ORANGE 44
#define CENTER_ALIGN -4108
#define LEFT_ALIGN -4131
//--
#define BOLD 0
#define ITALIC 1
#define UNDERLINE 2
//-
#define LINE_STYLE_CONTINUE 1
//XlBorderWeight
#define xlMedium -4138
#define xlThin 2
#define xlThick 4
#define xlHairline 1

#include <QObject>
#include <QAxObject>
class CWPSTable : public QObject {
	Q_OBJECT

public:
	CWPSTable() = default;
	//CWPSTable(QObject * parent = Q_NULLPTR);
	void getAddWorkBook();
	CWPSTable(bool visible, bool alert);//new a table
	void getOpenWorkBook(const QString& openPath);
	CWPSTable(const QString& openPath, bool visible, bool alert);//open a existed table
	~CWPSTable();
	bool newExcel();
	void deleteExcel();
	void save_Quit(const QString& path);
	void openWorkBooks(QString inOpenPath);
	void getWorkBooks();
	int countSheets();
	void getSheets();
	void addSheet(bool addAtLast = true);
	void copySheet();
	void CopyWorkBook();
	void addWorkBooks();
	void setVisible(bool inBool);
	void setAlert(bool inBool);
	void setCaption(QString inCaption);

	void activateCurrWorkBook();
	void getSheet(int inIndex);
	void setSheetName(const QString& name);
	QString getSheetName();
	void getRange(QString inRange);//EXAMPLE:"A1:A1","A1","A1:B3","A:A","3:3"...
	void getRange(int inRow1, int inColumn1);
	QAxObject* getRange(int inRow1, int inColumn1, int inRow2, int inColumn2);
	QString getValue(int inRow1, int inColumn1);
	QString getValue();
	void setValue(const QVariant& inValue);
	void setValue(int row, int col, const QVariant& value);
	void setComment(int row, int col,const QString& text);//enter="\n"
	void setRangeFormat(int row1, int col1, int row2, int col2, QString myFormat);//CELL RANGE FORMAT
	void autoFit();
	void autoFitColumn();
	void autoFitRow();
	void setBackColor(int inRow1, int inColumn1, int inRow2, int inColumn2, int inFontColor);
	void setFontColor(int inRow, int inColumn, int inFontColor);//EXAMPLE: 3 = RED
	void setFontColor(int inRow1, int inColumn1, int inRow2, int inColumn2, int inFontColor);//EXAMPLE: 3 = RED
	void copyPaste(int inRow1, int inColumn1, int inRow2, int inColumn2, int inPasteRow1, int inPasteColumn1, int inPasteRow2, int inPasteColumn2);
	void copyPasteFromSheet(int copySheetPos,int inRow1,int inColumn1,int inRow2,int inColumn2,int pasteSheetPos,int inPasteRow1,int inPasteColumn1);
	void copyInsertRowTo(int inRow1, int inColumn1, int inRow2, int inColumn2, int inPasteRow, int inPasteColumn, bool inIsDown, int inPasteRowNum);
	void insertRow(int inRow, int inColumn, bool inIsDown, int inRowNum);
	//	void getRange(int inRow1, int inColumn1, int inRow2, int inColumn2);
	//	void getCell(int inRow, int inColumn);//
//	QString getValue(QString inRange);

	//	void setValue(QString inRange, QString inValue);
		//	QString getCellValue(int inRow, int inColumn);
		//	void setCellValue(int inRow, int inColumn, QString inValue);
	void setScreenUpdating(bool inBool);
	void saveAs(QString inSavePath);//"C:\\Book1.xlsx"
	void quit();
	void setColumnWidth(int col, float width);
	void setColumnHidden(int col, bool isHidden = true);
	void mergeRange(int row1,int col1,int row2,int col2);
	void setRangeFontSize(int row1,int col1,int row2,int col2,int myFontSize);
	void setRangeAlign(int row1,int col1,int row2,int col2,int alignType);
	void setRangeFontStyle(int row1,int col1,int row2,int col2,int myStyle,bool styleFlag=true);
	void setRangeBorderStyle_Width(int row1,int col1,int row2,int col2,int borderStyle,int borderWidth= xlThin);
	void multiSelectCopy();

	QAxObject* MyExcel = nullptr;
	QAxObject* MyWorkBooks = nullptr;
	QAxObject* MyWorkBook = nullptr;
	QAxObject* mySheets = nullptr;
	QAxObject* mySheet = nullptr;
	QAxObject* range = nullptr;
	QAxObject* MyCell = nullptr;
};
