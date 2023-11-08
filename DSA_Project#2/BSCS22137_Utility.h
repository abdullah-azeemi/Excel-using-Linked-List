#pragma once
#include<windows.h>
#include<conio.h>
#include<iostream>
#include<vector>
#include<algorithm>
#include<iomanip>
#include<string>
#include<sstream>
#include<iterator>
#include<stack>
#include<fstream>
using namespace std;

static void getRowColbyLeftClick(int& rpos, int& cpos)
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
static void gotoRowCol(int rpos, int cpos)
{
	COORD scrn;
	HANDLE hOuput = GetStdHandle(STD_OUTPUT_HANDLE);
	scrn.X = cpos;
	scrn.Y = rpos;
	SetConsoleCursorPosition(hOuput, scrn);
}
static void SetClr(int clr)
{
	SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), clr);
}
static double calculateAverage(int sum, int numOfelements) {
	if (numOfelements == 0) {
		return 0.0;
	}

	double average = static_cast<double>(sum) / numOfelements;
	std::stringstream formattedAverage;
	formattedAverage << std::fixed << std::setprecision(2) << average;
	double result;
	formattedAverage >> result;
	return result;
}