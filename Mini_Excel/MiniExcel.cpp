#include<iostream>
#include "MiniExcel.h"
#include"Header1.h"
#include<conio.h>
#include<windows.h>
#include<string.h>
using namespace std;
int Cell_ColumnSize = 6;
int Cell_RowSize = 20;
int BoundryColor = 12;
int CurrentColor = 14;
int TextColor = 6;
int RangeSelectionColor = 14;
int Result = 5;
void color(int k)
{
	if (k > 255)
	{
		k = 1;
	}
	HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
	SetConsoleTextAttribute(hConsole, k);

}
void getRowColbyLeftClick(int& rpos, int& cpos)
{
	HANDLE hInput = GetStdHandle(STD_INPUT_HANDLE);
	DWORD Events;
	INPUT_RECORD InputRecord;
	SetConsoleMode(hInput, ENABLE_PROCESSED_INPUT | ENABLE_MOUSE_INPUT | ENABLE_EXTENDED_FLAGS);
	do
	{
		ReadConsoleInput(hInput, &InputRecord, 1, &Events);
		if (InputRecord.Event.MouseEvent.dwButtonState == FROM_LEFT_1ST_BUTTON_PRESSED)
		{
			cpos = InputRecord.Event.MouseEvent.dwMousePosition.X;
			rpos = InputRecord.Event.MouseEvent.dwMousePosition.Y;
			break;
		}
	} while (true);
}
void gotoRowCol(int rpos, int cpos)
{
	COORD scrn;
	HANDLE hOuput = GetStdHandle(STD_OUTPUT_HANDLE);
	scrn.X = cpos;
	scrn.Y = rpos;
	SetConsoleCursorPosition(hOuput, scrn);
}
MiniExcel::MiniExcel()
{
	NofRows = NofCols = 1;
	Head = Current = new Cell(0, 0);
	currentRow = currentCol = 0;
	int i = 0;
	while (i <3)
	{
		InsertColToRight(18);
		Current = Current->Right;
		currentCol++;
		i++;
	}
	i = 0;
	while (i<3)
	{
		InsertRowBelow(2);
		Current = Current->Down;
		currentRow++;
		i++;
	}
}
//Head
void MiniExcel::SetHead()
{
	MovePointerToLeft(Head);
	MovePointerToUp(Head);
}
//Current Functionalities
void MiniExcel::StoringActualCurrent()
{
	ActualCurrent = Current;

}
void MiniExcel::ReStoringActualCurrent()
{
	Current= ActualCurrent;
}
bool MiniExcel::IsMostAbove(Cell*Current)
{
	if (Current->Up==nullptr)
	{
		return true;
	}
	return false;
}
bool MiniExcel::IsMostRight(Cell* Current)
{
	if (Current->Right== nullptr)
	{
		return true;
	}
	return false;
}
bool MiniExcel::IsMostDown(Cell* Current)
{
	if (Current->Down== nullptr)
	{
		return true;
	}
	return false;
}
bool MiniExcel::IsMostLeft(Cell* Current)
{
	if (Current->Left== nullptr)
	{
		return true;
	}
	return false;
}
void MiniExcel::MoveCurrentToTop()
{
	while (Current!=nullptr and Current->Up!=nullptr)
	{
		Current = Current->Up;
	}
}
void MiniExcel::MoveCurrentToDown()
{
	while (Current->Down!= nullptr)
	{
		Current = Current->Down;

	}
}
void MiniExcel::MoveCurrentToLeft()
{
	while (Current!=nullptr and  Current->Left!= nullptr)
	{
		Current = Current->Left;
	}
}
void MiniExcel::MoveCurrentToRight()
{
	while (Current->Right!= nullptr)
	{
		Current = Current->Right;
	}
}
//Pointer Functions
void MiniExcel::MovePointerToRight(Cell*& Pointer)
{
	while (Pointer != nullptr and Pointer->Right!= nullptr)
	{
		Pointer = Pointer->Right;
	}
}
void MiniExcel::MovePointerToUp(Cell*& Pointer)
{
	while (Pointer != nullptr and Pointer->Up != nullptr)
	{
		Pointer = Pointer->Up;
	}
}
void MiniExcel::MovePointerToDown(Cell*& Pointer)
{
	while (Pointer != nullptr and Pointer->Down != nullptr)
	{
		Pointer = Pointer->Down;
	}
}
void MiniExcel::MovePointerToLeft(Cell*& Pointer)
{
	while (Pointer != nullptr and Pointer->Left!= nullptr)
	{
		Pointer = Pointer->Left;
	}
}
//User Interface
bool MiniExcel::ArrowKeys(char ch)
{
	if (ch == -32)
	{
		ch = _getch();
		if (ch==75)
		{
			//Left
			if (Current->Left)
			{
				PrintingCell(Current, BoundryColor,currentRow, currentCol);
				Current = Current->Left;
				currentCol--;
				PrintingCell(Current,CurrentColor,currentRow, currentCol);
			}
		}
		else if (ch == 77)
		{
			//Right
			if (Current->Right)
			{
				PrintingCell(Current, BoundryColor, currentRow, currentCol);
				Current = Current->Right;
				currentCol++;
				PrintingCell(Current, CurrentColor, currentRow, currentCol);
			}
		}
		else if (ch == 72)
		{
			//Up
			if (Current->Up)
			{
				PrintingCell(Current, BoundryColor, currentRow, currentCol);
				Current = Current->Up;
				currentRow--;
				PrintingCell(Current, CurrentColor, currentRow, currentCol);
			}
		}
		else if (ch == 80)
		{
			//Up
			if (Current->Down)
			{
				PrintingCell(Current, BoundryColor, currentRow, currentCol);
				Current = Current->Down;
				currentRow++;
				PrintingCell(Current,CurrentColor, currentRow, currentCol);
			}
		}
		return true;
	}
	return false;
}
bool MiniExcel::ArrowKeys2(char ch)
{
	if (ch == -32)
	{
		ch = _getch();
		if (ch == 75)
		{
			//Left
			if (Current->Left)
			{
				PrintingCell(Current, RangeSelectionColor, currentRow, currentCol);
				Current = Current->Left;
				currentCol--;
				PrintingCell(Current, CurrentColor, currentRow, currentCol);
			}
		}
		else if (ch == 77)
		{
			//Right
			if (Current->Right)
			{
				PrintingCell(Current, RangeSelectionColor, currentRow, currentCol);
				Current = Current->Right;
				currentCol++;
				PrintingCell(Current, CurrentColor, currentRow, currentCol);
			}
		}
		else if (ch == 72)
		{
			//Up
			if (Current->Up)
			{
				PrintingCell(Current, RangeSelectionColor, currentRow, currentCol);
				Current = Current->Up;
				currentRow--;
				PrintingCell(Current, CurrentColor, currentRow, currentCol);
			}
		}
		else if (ch == 80)
		{
			//Up
			if (Current->Down)
			{
				PrintingCell(Current, RangeSelectionColor, currentRow, currentCol);
				Current = Current->Down;
				currentRow++;
				PrintingCell(Current, CurrentColor, currentRow, currentCol);
			}
		}
		return true;
	}
	return false;
}
bool MiniExcel::InsertValue(char ch)
{
	
	if (ch == 6)
	{
		string x;
		int rpos1 = Current->row * Cell_ColumnSize - Current->row;
		int cCol1 = Current->col * Cell_RowSize - Current->col;
		int rpos2 = rpos1 + Cell_ColumnSize - 1;
		int cCol2 = cCol1 + Cell_RowSize - 1;
		gotoRowCol(12,Cell_RowSize*NofCols);
		cout << "Insert Value : ";
		cin >> x;
		Current->Data = x;
		gotoRowCol(12, Cell_RowSize * NofCols);
		cout << "                                                    \r";
		return true;
	}
	return false;

}
//Project Functions
bool MiniExcel::InsertColToRight(char ch)
{
	if (ch == 18)
	{
		Cell* Actual = Current;
		MoveCurrentToTop();
		Cell* Temp, * Previous = nullptr;
		while (Current != nullptr)
		{
			Temp = new Cell(currentRow, currentCol + 1);
			Temp->Left = Current;
			Temp->Right = Current->Right;
			if (Current->Right)
				Current->Right->Left = Temp;
			Current->Right = Temp;
			Temp->Up = Previous;
			if (Previous)
				Previous->Down = Temp;
			if (Current->Down != nullptr)
				Current = Current->Down;
			else
				break;
			Previous = Temp;
		}
		NofCols++;
		SetHead();
		Current = Actual;
	
		return true;
	}
	return false;
}
bool MiniExcel::InsertColToLeft(char ch)
{

	if (ch==12)
	{
		Cell* Actual = Current;
		MoveCurrentToTop();
		Cell* Temp, * Previous = nullptr;
		while (Current != nullptr)
		{
			Temp = new Cell(currentRow, currentCol - 1);
			Temp->Right = Current;
			Temp->Left = Current->Left;
			if (Current->Left)
				Current->Left->Right = Temp;
			Current->Left = Temp;
			Temp->Up = Previous;
			if (Previous)
				Previous->Down = Temp;
			if (Current->Down != nullptr)
				Current = Current->Down;
			else
				break;
			Previous = Temp;
		}
		Current = Actual;
		currentCol++;
		NofCols++;
		SetHead();
		return true;
	}
	return false;
}
bool MiniExcel::InsertRowBelow(char ch)
{
	if (ch==2)
	{
		Cell* Actual = Current;
		MoveCurrentToLeft();
		Cell* Temp, * Below, * Previous = nullptr, * After = nullptr;
		while (Current != nullptr)
		{
			Temp = new Cell(currentRow + 1, currentCol);
			Temp->Down = Current->Down;
			Current->Down = Temp;
			Below = Temp->Down;
			if (Below != nullptr)
				Below->Up = Temp;
			Temp->Up = Current;
			Temp->Left = Previous;
			if (Previous != nullptr)
				Previous->Right = Temp;
			Temp->Right = After;
			if (Current->Right != nullptr)
			{
				Current = Current->Right;
			}
			else
				break;
			Previous = Temp;
		}
		NofRows++;
		SetHead();
		Current = Actual;
		return true;
	}
	return false;
}
bool MiniExcel::InsertRowAbove(char ch)
{
	if (ch == 1)
	{
		Cell* Actual = Current;
		MoveCurrentToLeft();
		Cell* Temp, * Below, * Previous = nullptr, * After = nullptr;
		while (Current != nullptr)
		{
			Temp = new Cell(currentRow, currentCol);
			Temp->Up = Current->Up;
			Current->Up = Temp;
			Below = Temp->Up;
			if (Below != nullptr)
				Below->Down = Temp;
			Temp->Down = Current;
			Temp->Left = Previous;
			if (Previous != nullptr)
				Previous->Right = Temp;
			Temp->Right = After;
			if (Current->Right != nullptr)
			{
				Current = Current->Right;
				currentCol++;
			}
			else
				break;
			Previous = Temp;
		}
		NofRows++;
		SetHead();
		Current = Actual;
		return true;
	}
	return false;
}
bool MiniExcel::InsertCellByRightShift(char ch)
{
	if (ch==17)
	{
		Cell* Actual = Current;
		MoveCurrentToRight();
		InsertColToRight(18);
		MoveCurrentToRight();
		while (1)
		{
			Current->Data = Current->Left->Data;
			if (Current->Left == Actual)
			{
				Current->Left->Data = "";
				break;
			}
			Current = Current->Left;
		}
		Current = Actual;

		return true;
	}
	return false;
}
bool MiniExcel::InsertCellByDownShift(char ch)
{
	if (ch == 23)
	{
		Cell* Actual = Current;
		MoveCurrentToDown();
		InsertRowBelow(2);
		MoveCurrentToDown();
		while (1)
		{
			Current->Data = Current->Up->Data;
			if (Current->Up == Actual)
			{
				Current->Up->Data = "";
				break;
			}
			Current = Current->Up;
		}
		Current = Actual;

		return true;
	}
	return false;
}
bool MiniExcel::DeleteCellByLeftShift(char ch)
{
	Cell* Actual = Current;
	if (ch==15)
	{
		while (Current->Right)
		{
			Current->Data = Current->Right->Data;
			Current->Right->Data = "";
			if (Current->Right)
				Current = Current->Right;
			else
				break;
		}

		Current = Actual;
		return true;
	}
	return false;
}
bool MiniExcel::DeleteCellByUpShift(char ch)
{
	Cell* Actual = Current;
	if (ch == 16)
	{
		while (Current->Down)
		{
			Current->Data = Current->Down->Data;
			Current->Down->Data = "";
			if (Current->Down)
				Current = Current->Down;
			else
				break;
		}
		Current = Actual;
		return true;
	}
	return false;
}
bool MiniExcel::DeleteColumn(char ch)
{
	if (ch==20)
	{
	
		Cell* Left;
		Cell* Right;
		Cell* Temp=Current;
		if (currentCol == NofCols - 1)
		{
			Current = Current->Left;
			currentCol--;
		}
		else if(currentCol == 0)
		{
			Current = Current->Right;
		}
		MovePointerToUp(Temp);
		while (true)
		{
			Left= Temp->Left;
			Right= Temp->Right;
			if(Left)
				Left->Right = Right;
			if(Right)
				Right->Left = Left;
			if (Temp->Down)
			{
				Temp = Temp->Down;
			}
			else
				break;
		}
		SetHead();
		NofCols--;
		return true;
	}
	return false;
}
bool MiniExcel::DeleteRow(char ch)
{
	if (ch == 5)
	{
		Cell* Up;
		Cell* Down;
		Cell* Temp = Current;
		if (currentRow== NofRows- 1)
		{
			Current = Current->Up;
			currentRow--;
		}
		else if (currentRow== 0)
		{
			Current = Current->Down;
		}
		MovePointerToLeft(Temp);
		while (true)
		{
			Up = Temp->Up;
			Down = Temp->Down;
			if (Up)
				Up->Down= Down;
			if (Down)
				Down->Up= Up;
			if (Temp->Right)
			{
				Temp = Temp->Right;
			}
			else
				break;
		}
		SetHead();
		NofRows--;
		return true;
	}
	return false;
}
bool MiniExcel::ClearColumn(char ch)
{
	if (ch == 8)
	{
		Cell* Actual = Current;
		MoveCurrentToTop();
		while (true)
		{
			Current->Data = "";
			if (Current->Down)
				Current = Current->Down;
			else
				break;
		}
		Current = Actual;


		return true;
	}
	return false;
}
bool MiniExcel::ClearRow(char ch)
{
	if (ch == 21)
	{
		Cell* Actual = Current;
		MoveCurrentToLeft();
		while (true)
		{
			Current->Data = "";
			if (Current->Right)
				Current = Current->Right;
			else
				break;
		}
		Current = Actual;

		return true;
	}
	return false;
}
bool MiniExcel::GetRangeSum(char ch)
{
	if (ch == 19)
	{
		return true;
	}
	return false;


}
void MiniExcel::SwapingPointers(Cell*&Start,Cell*End)
{
	Cell* Temp = Start;
	Start = End;
	End = Temp;
}
bool MiniExcel::Copy(char ch)
{
	if (ch == 3)
	{
		vector<string>Temp;
		COPY = Temp;

		Cell* Actual = Current;
		int arow = currentRow;

		int acol = currentCol;
		Cell* Starting = nullptr;
		Cell* Ending = nullptr;
		SelectRange(Starting, Ending, Actual, arow, acol);

		if (IsStartLeftTORange(Starting, Ending))
		{
			while (Starting != Ending)
			{
				COPY.push_back(Starting->Data);
				Starting = Starting->Right;
			}
			COPY.push_back(Starting->Data);
			PasteWise = "row";
		}
		if (IsStartRightToRange(Starting, Ending))
		{
			SwapingPointers(Starting, Ending);

			while (Starting != Ending)
			{
				COPY.push_back(Starting->Data);
				Starting = Starting->Right;
			}
			COPY.push_back(Starting->Data);
			PasteWise = "row";
		}
		else if (IsStartUpToRange(Starting, Ending))
		{
			while (Starting != Ending)
			{
				COPY.push_back(Starting->Data);
				Starting = Starting->Down;
			}
			COPY.push_back(Starting->Data);
			PasteWise = "col";
		}
		else if (IsStartDownToRange(Starting, Ending))
		{

			SwapingPointers(Starting, Ending);
			while (Starting != Ending)
			{
				COPY.push_back(Starting->Data);
				Starting = Starting->Down;
			}
			COPY.push_back(Starting->Data);
			PasteWise = "col";

		}
		return true;
	}
	return false;

}
bool MiniExcel::PASTE(char ch)
{
	bool CellSelected = false;
	Cell* Previous = nullptr;
	Cell* PasteHead = nullptr;
	if (ch==22)
	{
		gotoRowCol(12, Cell_RowSize * NofCols);
		cout << "                               ";
		gotoRowCol(12, Cell_RowSize * NofCols);
		cout << "Please select Cell  ";

		ch = _getch();
		if (ArrowKeys(ch))
		{
		}
		while (!CellSelected)
		{
			ch = _getch();

			if (ch == 19)
			{
				PasteHead = Current;
				CellSelected = true;
			}
			else if (ArrowKeys(ch))
			{
			}

			if (ch==19)
			{
				PasteHead = Current;
				CellSelected = true;
			}
		}
		gotoRowCol(12, Cell_RowSize * NofCols);
		cout << "                                 ";
		int temp = 0;
		if (PasteWise == "row")
		{
			while (temp!=COPY.size())
			{
				if (!PasteHead)
				{
					PasteHead = Previous;
					Cell* X = Current;
					Current = PasteHead;
					InsertColToRight(18);
					Current = X;
					PasteHead = PasteHead->Right;
				}

				PasteHead->Data = COPY[temp];
				temp++;
				Previous = PasteHead;
				PasteHead = PasteHead->Right;

			}
		}
		else if (PasteWise == "col")
		{
			while (temp != COPY.size())
			{
				if (!PasteHead)
				{
					PasteHead = Previous;
					Cell* X = Current;
					Current = PasteHead;
					InsertRowBelow(2);
					Current = X;
					PasteHead = PasteHead->Down;
				}

				PasteHead->Data = COPY[temp];
				temp++;
				Previous = PasteHead;
				PasteHead = PasteHead->Down;

			}

		}

		return true;
	}

	return false;

}
bool MiniExcel::CUT(char ch)
{

	if (ch == 24)
	{
		vector<string>Temp;
		COPY = Temp;

		Cell* Actual = Current;
		int arow = currentRow;

		int acol = currentCol;
		Cell* Starting = nullptr;
		Cell* Ending = nullptr;
		SelectRange(Starting, Ending, Actual, arow, acol);

		if (IsStartLeftTORange(Starting, Ending))
		{
			while (Starting != Ending)
			{
				COPY.push_back(Starting->Data);
				Starting->Data = "";

				Starting = Starting->Right;
			}
			COPY.push_back(Starting->Data);
			Starting->Data = "";
			PasteWise = "row";
		}
		if (IsStartRightToRange(Starting, Ending))
		{
			SwapingPointers(Starting, Ending);

			while (Starting != Ending)
			{
				COPY.push_back(Starting->Data);
				Starting->Data = "";

				Starting = Starting->Right;
			}
			COPY.push_back(Starting->Data);
			Starting->Data = "";

			PasteWise = "row";
		}
		else if (IsStartUpToRange(Starting, Ending))
		{
			while (Starting != Ending)
			{
				COPY.push_back(Starting->Data);
				Starting->Data = "";

				Starting = Starting->Down;
			}
			COPY.push_back(Starting->Data);
			Starting->Data = "";

			PasteWise = "col";
		}
		else if (IsStartDownToRange(Starting, Ending))
		{

			SwapingPointers(Starting, Ending);
			while (Starting != Ending)
			{
				COPY.push_back(Starting->Data);
				Starting->Data = "";

				Starting = Starting->Down;
			}
			COPY.push_back(Starting->Data);
			Starting->Data = "";

			PasteWise = "col";

		}
		return true;
	}
	return false;



}
//Range
bool MiniExcel::IsStartLeftTORange(Cell* Start, Cell* End)
{
	while (true)
	{
		if (Start->Right)
			Start = Start->Right;
		else
			return false;
		if (Start == End)
			return true;

	}
	return false;
}
bool MiniExcel::IsStartRightToRange(Cell* Start, Cell* End)
{
	while (true)
	{
		if (Start->Left)
			Start = Start->Left;
		else
			return false;
		if (Start == End)
			return true;
	}
	return false;
}
bool MiniExcel::IsStartUpToRange(Cell* Start, Cell* End)
{
	while (true)
	{
		if (Start->Down)
			Start = Start->Down;
		else
			return false;
		if (Start == End)
			return true;
	}
	return false;
}
bool MiniExcel::IsStartDownToRange(Cell* Start, Cell* End)
{
	while (true)
	{
		if (Start->Up)
			Start = Start->Up;
		else
			return false;
		if (Start == End)
			return true;
	}
	return false;
}
bool MiniExcel::isNumber(const string& str)
{
	for (char const& c : str) {
		if (isdigit(c) == 0) return false;
	}
	return true;
}
void MiniExcel::SelectRange(Cell*& Initial, Cell*& Final,Cell*&Actual,int arow,int acol)
{

	bool StartSelected = false;
	bool RangeSelected = false;
	char ch;
	gotoRowCol(12, Cell_RowSize * NofCols);
	cout << "Please select Range Starting : ";
	while (!StartSelected)
	{
		ch = _getch();
		if (ArrowKeys(ch))
		{
			Actual->Data = "  ";
			PrintingCell(Current, Result, arow, acol);
		}
		else if (ch == 19)
		{
			Initial = Current;
			StartSelected = true;
			gotoRowCol(12, Cell_RowSize * NofCols);
			cout << "Please select Range Ending : ";
			while (!RangeSelected)
			{
				ch = _getch();
				if (ArrowKeys2(ch))
				{
				}
				else if (ch == 19)
				{
					Final= Current;
					RangeSelected = true;
				}
			}
		}
	}

}

//Printing Functions
bool MiniExcel::IsFunctions(char ch)
{
	int row = 0, col = 0;
	if (ch==25)
	{
		Cell* Actual = Current, * Ending = nullptr, * Starting = nullptr;
		int arow = currentRow;
		int acol= currentCol;
		PrintingCell(Current, Result, arow, acol);

		SelectRange(Starting, Ending,Actual,arow,acol);

		gotoRowCol(12, Cell_RowSize * NofCols);
		cout << "Please select the function ";
		getRowColbyLeftClick(row,col);
		row = row / Cell_ColumnSize;
		col = col / Cell_RowSize;
		gotoRowCol(12, Cell_RowSize * NofCols);
		cout << "                                        ";
		int sum = 0, value = 0,count=1,average=0,Minimum=0,Maximum=0, COUNT=0;
		Current = Starting;
		if (row==3 and (col==NofCols+4 or col==NofCols+3))
		{
			if (IsStartLeftTORange(Starting, Ending))
			{
				while (Current != Ending)
				{
					value = stoi(Current->Data);
					sum += value;
					Current = Current->Right;
				}
				value = stoi(Current->Data);
				sum += value;
			}
			else if (IsStartRightToRange(Starting, Ending))
			{

				while (Current != Ending)
				{
					value = stoi(Current->Data);
					sum += value;
					Current = Current->Left;
				}
				value = stoi(Current->Data);
				sum += value;


			}
			else if (IsStartUpToRange(Starting, Ending))
			{

				while (Current != Ending)
				{
					value = stoi(Current->Data);
					sum += value;
					Current = Current->Down;
				}
				value = stoi(Current->Data);
				sum += value;


			}
			else if (IsStartDownToRange(Starting, Ending))
			{

				while (Current != Ending)
				{
					value = stoi(Current->Data);
					sum += value;
					Current = Current->Up;
				}
				value = stoi(Current->Data);
				sum += value;

			}
		}
		else if (row==2 and (col == NofCols + 4 or col == NofCols + 3))
		{
			if (IsStartLeftTORange(Starting, Ending))
			{
				while (Current != Ending)
				{
					value = stoi(Current->Data);
					sum += value;
					count++;
					Current = Current->Right;
				}
				value = stoi(Current->Data);
				sum += value;
			}
			else if (IsStartRightToRange(Starting, Ending))
			{

				while (Current != Ending)
				{
					value = stoi(Current->Data);
					sum += value;
					count++;

					Current = Current->Left;
				}
				value = stoi(Current->Data);
				sum += value;


			}
			else if (IsStartUpToRange(Starting, Ending))
			{

				while (Current != Ending)
				{
					value = stoi(Current->Data);
					sum += value;
					count++;

					Current = Current->Down;
				}
				value = stoi(Current->Data);
				sum += value;


			}
			else if (IsStartDownToRange(Starting, Ending))
			{

				while (Current != Ending)
				{
					value = stoi(Current->Data);
					sum += value;
					count++;

					Current = Current->Up;
				}
				value = stoi(Current->Data);
				sum += value;
			}
			average = sum / count;
			sum = average;
		}		
		else if (row == 0 and (col == NofCols + 4 or col == NofCols + 3))
		{

			if (IsStartLeftTORange(Starting, Ending))
			{

				value = stoi(Current->Data);
				Minimum = value;
				Current->Right;
				while (Current != Ending)
				{
					value = stoi(Current->Data);
					if (value < Minimum)
						Minimum = value;
					Current = Current->Right;
				}
				value = stoi(Current->Data);
				if (value < Minimum)
					Minimum = value;
			}
			else if (IsStartRightToRange(Starting, Ending))
			{

				value = stoi(Current->Data);
				Minimum = value;
				Current->Left;
				while (Current != Ending)
				{
					value = stoi(Current->Data);
					if (value < Minimum)
						Minimum = value;
					Current = Current->Left;
				}
				value = stoi(Current->Data);
				if (value < Minimum)
					Minimum = value;

			}
			else if (IsStartUpToRange(Starting, Ending))
			{

				value = stoi(Current->Data);
				Minimum = value;
				Current->Down;
				while (Current != Ending)
				{
					value = stoi(Current->Data);
					if (value < Minimum)
						Minimum = value;
					Current = Current->Down;
				}
				value = stoi(Current->Data);
				if (value < Minimum)
					Minimum = value;


			}
			else if (IsStartDownToRange(Starting, Ending))
			{

				value = stoi(Current->Data);
				Minimum = value;
				Current->Up;
				while (Current != Ending)
				{
					value = stoi(Current->Data);
					if (value < Minimum)
						Minimum = value;
					Current = Current->Up;
				}
				value = stoi(Current->Data);
				if (value < Minimum)
					Minimum = value;
			}
			sum = Minimum;
		}
		else if (row == 5 and (col == NofCols + 4 or col == NofCols + 3))
		{

			if (IsStartLeftTORange(Starting, Ending))
			{
	
				value = stoi(Current->Data);
				Maximum = value;
				Current->Right;
				while (Current != Ending)
				{
					value = stoi(Current->Data);
					if (value > Maximum)
						Maximum = value;
					Current = Current->Right;
				}
				value = stoi(Current->Data);
				if (value > Maximum)
					Maximum = value;
			}
			else if (IsStartRightToRange(Starting, Ending))
			{
	
				value = stoi(Current->Data);
				Maximum = value;
				Current->Left;
				while (Current != Ending)
				{
					value = stoi(Current->Data);
					if (value > Maximum)
						Maximum = value;
					Current = Current->Left;
				}
				value = stoi(Current->Data);
				if (value > Maximum)
					Maximum = value;

			}
			else if (IsStartUpToRange(Starting, Ending))
			{
	
				value = stoi(Current->Data);
				Maximum = value;
				Current->Down;
				while (Current != Ending)
				{
					value = stoi(Current->Data);
					if (value > Maximum)
					Maximum = value;
					Current = Current->Down;
				}
				value = stoi(Current->Data);
				if (value > Minimum)
					Maximum = value;


			}
			else if (IsStartDownToRange(Starting, Ending))
			{
	
				value = stoi(Current->Data);
				Maximum = value;
				Current->Up;
				while (Current != Ending)
				{
					value = stoi(Current->Data);
					if (value > Maximum)
						Maximum = value;
					Current = Current->Up;
				}
				value = stoi(Current->Data);
				if (value > Maximum)
					Maximum = value;
			}
			sum = Maximum;
		}
		else if (row == 7 and (col == NofCols + 4 or col == NofCols + 3))
		{

			if (IsStartLeftTORange(Starting, Ending))
			{

				while (Current != Ending)
				{
					if (isNumber(Current->Data))
					{
						COUNT++;
					}
					Current = Current->Right;
				}
				if (isNumber(Current->Data))
				{
					COUNT++;
				}
			}
			else if (IsStartRightToRange(Starting, Ending))
			{

				while (Current != Ending)
				{
					if (isNumber(Current->Data))
					{
						COUNT++;
					}
					Current = Current->Left;
				}
				if (isNumber(Current->Data))
				{
					COUNT++;
				}
			}
			else if (IsStartUpToRange(Starting, Ending))
			{

				while (Current != Ending)
				{
					if (isNumber(Current->Data))
					{
						COUNT++;
					}
					Current = Current->Down;
				}
				if (isNumber(Current->Data))
				{
					COUNT++;
				}
			}
			else if (IsStartDownToRange(Starting, Ending))
			{

				while (Current != Ending)
				{
					if (isNumber(Current->Data))
					{
						COUNT++;
					}
					Current = Current->Down;
				}
				if (isNumber(Current->Data))
				{
					COUNT++;
				}
			}
		    sum = COUNT;
		}
		Actual->Data = to_string(sum);
		Current = Actual;
		currentRow = arow;
		currentCol= acol;
		return true;
	}
	return false;
}
void MiniExcel::PrintingFunctions()
{

	Cell* Maximum= new Cell(0, 0, "Maximum");
	PrintingCell(Maximum, BoundryColor, 6, NofCols + 4);

	Cell* Minimum = new Cell(0, 0, "Minimum");
	PrintingCell(Minimum, BoundryColor, 0, NofCols + 4);

	Cell* Average = new Cell(0,0,"Average");
	PrintingCell(Average,BoundryColor,2,NofCols+4);

	Cell* Sum = new Cell(0, 0, "SUM");
	PrintingCell(Sum,BoundryColor,4,NofCols+4);


	Cell* COUNT= new Cell(0, 0, "Count");
	PrintingCell(COUNT, BoundryColor, 8, NofCols + 4);

}
void MiniExcel::Print()
{

	MovePointerToLeft(Head);
	MovePointerToUp(Head);
	cout << currentRow<<"," << currentCol << endl;
	cout << NofRows << ":" << NofCols << endl;
	while (Head!=nullptr)
	{

		while (Head != nullptr)
		{
			cout << Head << "=>"<<"("<<Head->row<<","<<Head->col<<")";
			cout << Head->Up << "," << Head->Left << "," << Head->Right << "," << Head->Down;
			cout << "         ";
			if (Head->Right != nullptr)
				Head = Head->Right;
			else
				break;
		}
		MovePointerToLeft(Head);
		if (Head->Down != nullptr)
			Head = Head->Down;
		else
			break;
		cout << endl;
	}
	MovePointerToLeft(Head);
	MovePointerToUp(Head);

}
void MiniExcel::PrintingGrid()
{
	PrintingFunctions();
	color(BoundryColor);
	MovePointerToUp(Head);
	MovePointerToLeft(Head);
	for (int r=0;r<NofRows;r++)
	{
		for (int c = 0; c < NofCols; c++)
		{
			PrintingCell(Head,BoundryColor,r,c);
			if(Head->Right)
				Head = Head->Right;
			cout << endl;
		}
		if (Head->Down)
		{
			Head = Head->Down;
			MovePointerToLeft(Head);
		}
		
	}
	PrintingCell(Current,CurrentColor,currentRow,currentCol);

	MovePointerToUp(Head);
	MovePointerToLeft(Head);
	gotoRowCol(NofCols * Cell_ColumnSize,0);
}
void MiniExcel::PrintingCell(Cell* Temp,int col,int r,int c)
{
	color(col);
	int rpos1 = r * Cell_ColumnSize -r;
	int cCol1 = c * Cell_RowSize -c;
	int rpos2 = rpos1+ Cell_ColumnSize -1;
	int cCol2 = cCol1+ Cell_RowSize -1;
	for (int ri=rpos1;ri<=rpos2;ri++)
	{
		gotoRowCol(ri, cCol1);
		cout << char(-37);
		gotoRowCol(ri, cCol2);
		cout << char(-37);
	}
	gotoRowCol(((rpos1 + rpos2) / 2)+1, ((cCol1 + cCol2) / 2)-1);
	color(TextColor);
	cout << "         ";
	gotoRowCol(((rpos1 + rpos2) / 2) + 1, ((cCol1 + cCol2) / 2) - 1);
	cout << Temp->Data;
	color(col);
	for (int ci = cCol1; ci <= cCol2; ci++)
	{
		gotoRowCol(rpos1, ci);
		cout << char(-37);
		gotoRowCol(rpos2, ci);
		cout << char(-37);
	}

}
