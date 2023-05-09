#pragma once
#include<string>
#include<vector>
#include<Windows.h>
#include<iostream>
using namespace std;
struct ActualPosition
{
	int row;
	int Col;
};
class MiniExcel
{
private:
	class Cell
	{
	private:
		friend class MiniExcel;
		Cell* Right, * Left, * Up, * Down;
		string Data;
		int row, col;
	public:
		Cell(int Ro = 0, int C = 0, string D = "", Cell* R = nullptr, Cell* L = nullptr, Cell* U = nullptr, Cell* Dn = nullptr)
		{
			row = Ro;
			col = C;
			Data = D;
			Right = R;
			Left = L;
			Up = U;
			Down = Dn;
		}
	};
	Cell* ActualCurrent;
	Cell* Head;
	Cell* Tail;
	Cell* Current;
	vector<string> COPY;
	string PasteWise;
	int currentRow;
	int currentCol;
	int NofRows;
	int NofCols;
public:
	MiniExcel();
	//Current
	bool IsMostAbove(Cell* Current);
	bool IsMostRight(Cell* Current);
	bool IsMostDown(Cell* Current);
	bool IsMostLeft(Cell* Current);
	void MoveCurrentToTop();
	void MoveCurrentToDown();
	void MoveCurrentToLeft();
	void MoveCurrentToRight();
	void StoringActualCurrent();
	void ReStoringActualCurrent();
	//MainFunctions
	bool InsertColToRight(char ch);
	bool InsertColToLeft(char ch);
	bool InsertRowBelow(char ch);
	bool InsertRowAbove(char ch);
	bool InsertCellByRightShift(char ch);
	bool InsertCellByDownShift(char ch);
	bool DeleteCellByLeftShift(char ch);
	bool DeleteCellByUpShift(char ch);
	bool DeleteColumn(char ch);
	bool DeleteRow(char ch);
	bool ClearColumn(char ch);
	bool ClearRow(char ch);
	bool GetRangeSum(char ch);
	bool Copy(char ch);
	bool PASTE(char ch);
	bool CUT(char ch);

	//Functions
	void PrintingFunctions();
	bool IsFunctions(char ch);
	void PrintingGrid();
	void Print();
	//Cells
	void PrintingCell(Cell* Temp, int color, int r, int c);
	//Pointer
	void SetHead();
	void MovePointerToUp(Cell*& Pointer);
	void MovePointerToDown(Cell*& Pointer);
	void MovePointerToLeft(Cell*& Pointer);
	void MovePointerToRight(Cell*& Pointer);
	void SwapingPointers(Cell*& Start, Cell* End);

	//Keys
	bool ArrowKeys(char ch);
	bool ArrowKeys2(char ch);
	bool InsertValue(char ch);
	//Range
	bool IsStartLeftTORange(Cell* Start, Cell* End);
	bool IsStartRightToRange(Cell* Start, Cell* End);
	bool IsStartUpToRange(Cell* Start, Cell* End);
	bool IsStartDownToRange(Cell* Start, Cell* End);
	bool isNumber(const string& str);
	void SelectRange(Cell*& Initial, Cell*& Final, Cell*& Actual, int arow, int acol);



};
