#include<iostream>
#include"MiniExcel.h"
#include<conio.h>
#include"Header1.h"
using namespace std;

int main()
{
	MiniExcel E;
	bool ToDisplay = false;
	bool ToClear= false;
	E.PrintingGrid();

	cout << endl;
	while (1)
	{
		ToDisplay = false;
		ToClear= false;
		char ch = _getch();
		if(E.ArrowKeys(ch))
		{
			ToDisplay = false;
		}
		else if(E.InsertValue(ch))
		{
			ToDisplay = false;
		}
		else if(E.InsertCellByRightShift(ch))
		{

			ToDisplay = true;
		}
		else if(E.InsertColToRight(ch))
		{
			ToDisplay = true;
		}
		else if(E.InsertColToLeft(ch))
		{
			ToDisplay = true;
		}
		else if(E.InsertRowAbove(ch))
		{
			ToDisplay = true;
		}
		else if(E.InsertRowBelow(ch))
		{
			ToDisplay = true;
		}
		else if(E.InsertCellByDownShift(ch))
		{
			ToDisplay = true;
		}
		else if(E.DeleteCellByLeftShift(ch))
		{
			ToDisplay = true;
		}
		else if(E.DeleteCellByUpShift(ch))
		{
			ToDisplay = true;
		}
		else if(E.DeleteColumn(ch))
		{
			ToClear = true;

			ToDisplay = true;
		}
		else if (E.DeleteRow(ch)) 
		{
			//Problem
			ToClear = true;
			ToDisplay = true;
		}
		else if (E.ClearColumn(ch)) 
		{
			ToDisplay = true;
		}
		else if (E.ClearRow(ch)) 
		{
			ToDisplay = true;
		}
		else if (E.IsFunctions(ch))
		{
			ToDisplay = true;
		}
		else if (E.Copy(ch))
		{
			ToDisplay = true;
			ToClear = false;
		}
		else if (E.PASTE(ch))
		{
			ToDisplay = true;
			ToClear = true;
		}
		else if (E.CUT(ch))
		{
			ToDisplay = true;
		}
		if (ToClear)
		{
			system("cls");
		}
		if (ToDisplay)
		{
			E.PrintingGrid();
		}
	}
}
