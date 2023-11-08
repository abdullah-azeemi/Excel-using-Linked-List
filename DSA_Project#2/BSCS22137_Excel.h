#pragma once
#include"BSCS22137_Utility.h"
class Excel
{
	int cellWidth, cellHeight;

	class Cell {

		string data;
		Cell* right;
		Cell* left;
		Cell* up;
		Cell* down;
		friend class Excel;

	public:
		Cell(char data = ' ', Cell* r = nullptr, Cell* l = nullptr, Cell* u = nullptr, Cell* d = nullptr) {
			data = data;
			right = r;
			left = l;
			up = u;
			down = d;
		}
	};
	
	int rSize = 0, cSize = 0;
	int rCurr = 0, cCurr = 0;
	Cell* curr;
	Cell* head; 
	Cell* tail;
	Cell* rangeStart = nullptr;
	Cell* rangeEnd = nullptr;
	int rs, cs, re, ce;
	bool rangeS,rangeE;
	vector<string> store;

public:

	friend class Cell;
	Excel() {
		cellWidth = 8;
		cellHeight = 4;
		rSize = 0;
		cSize = 0;
		head = new Cell();
		tail = head;
		curr = head;
		rCurr = 0;
		cCurr = 0;

		rangeS = false, rangeE = false;

		for (int ri = 1; ri < 5; ri++) {
			insertColRight();
			curr = curr->right;
		}
		curr = head;
		for (int ri = 1; ri < 5; ri++) {
			insertrowDown();
			curr = curr->down;
		}
		curr = head;
		printGrid();
		//printData();
		printCellBorder(rCurr, cCurr, 4);
		printButtons();
		inputKeyboard();
	}
	void insertColRight() {

		Cell* temp = curr;
		while (curr->up != nullptr) {
			curr = curr->up;
		}
		while (curr != nullptr) {
			insertCellRight(curr);
			curr = curr->down; 
		}
		curr = temp;
		cSize++;
	}
	void insertrowDown() {
		Cell* temp = curr;
		while (curr->left != nullptr) {
			curr = curr->left;
		}
		while (curr != nullptr) {
			insertCellDown(curr);
			curr = curr->right;
		}
		curr = temp;
		rSize++;
	}
	void insertRowUp() {
		Cell* temp = curr;
		while (curr->left != nullptr) {
			curr = curr->left;
		}
		while (curr != nullptr) {
			insertCellUp(curr);
			curr = curr->right;
		}
		curr = temp;
		rSize++;
	}
	void insertColLeft() {
		Cell* temp = curr;
		while (curr->up != nullptr) {
			curr = curr->up;
		}
		while (curr != nullptr) {
			insertCellLeft(curr);
			curr = curr->down;
		}
		curr = temp;
		cSize++;
	}


	void deleteCol() {
		while (curr->up != nullptr) {
			curr = curr->up;
		}
		if (curr->left == nullptr) {
			curr = curr->right;
			while (curr->down != nullptr) {
				curr->left = nullptr;
				curr = curr->down;
			}
			cCurr++;
		}
		else {
			curr = curr->left;
			while (curr->down != nullptr) {
				curr->right = nullptr;
				curr = curr->down;
			}
			cCurr--;
		}
		cSize--;
	
	}
	void deleteRow() {
		while (curr->left != nullptr) {
			curr = curr->left;
		}
		if (curr->up == nullptr) {
			curr = curr->down;
			while (curr->right != nullptr) {
				curr->up = nullptr;
				curr = curr->right;
			}
			rCurr++;
		}
		else {
			curr = curr->up;
			while (curr->right != nullptr) {
				curr->down = nullptr;
				curr = curr->right;
			}
			rCurr--;
		}
		rSize--;

	}

	Cell* insertCellRight(Cell*& c) {
		Cell* temp = new Cell();
		temp->left = c;
		if (c->right != nullptr) {
			temp->right = c->right;
			temp->right->left = temp;
		}
		c->right = temp;
		if (c->up != nullptr && c->up->right != nullptr) {
			temp->up = c->up->right;
			c->up->right->down = temp;
		}
		if (c->down != nullptr && c->down->right != nullptr) {
			temp->down = c->down->right;
			c->down->right->up = temp;
		}
		return temp;
	}
	Cell* insertCellDown(Cell*& c) {
		Cell* temp = new Cell();
		temp->up = c;
		if (c->down != nullptr) {
			temp->down = c->down;
			temp->down->up = temp;
		}
		c->down = temp;
		if (c->left != nullptr && c->left->down != nullptr) {
			temp->left = c->left->down;
			c->left->down->right = temp;
		}
		if (c->right != nullptr && c->right->down != nullptr) {
			temp->right = c->right->down;
			c->right->down->left = temp;
		}
		return temp;
	}
	Cell* insertCellUp(Cell*& c) {
		Cell* temp = new Cell();
		temp->down = c;
		if (c->up != nullptr) {
			temp->up = c->up;
			temp->up->down = temp;
		}
		c->up = temp;
		if (c->right != nullptr && c->right->up != nullptr) {
			temp->right = c->right->up;
			c->right->up->left = temp;
		}
		if (c->left != nullptr && c->left->up != nullptr) {
			temp->left = c->left->up;
			c->left->up->right = temp;
		}
		return temp;
	}
	Cell* insertCellLeft(Cell*& c) {
		Cell* temp = new Cell();
		temp->right = c;
		if (c->left != nullptr) {
			temp->left = c->left;
			temp->left->right = temp;
		}
		c->left = temp;
		if (c->up != nullptr && c->up->left != nullptr) {
			temp->up = c->up->left;
			c->up->left->down = temp;
		}
		if (c->down != nullptr && c->down->left != nullptr) {
			temp->down = c->down->left;
			c->down->left->up = temp;
		}
		return temp;
	}

	

	void printGrid() {
		int rIn = 0, cIn = 0;
		for (int ri = 0; ri < rSize; ri++) {
			for (int ci = 0; ci < cSize; ci++) {
				printCellBorder(rIn,cIn);
				cIn += 8;
			}
			rIn += 3;
			cIn = 0;
		}
	}
	void printCell(int rri, int cci, int clr = 7) {
		SetClr(clr);
		int cellWidth = 12, cellHeight = 4;
		for (int ri = 0; ri < cellWidth; ri++) {
			for (int ci = 0; ci < cellHeight; ci++) {
				gotoRowCol(ri + rri, ci + cci);
				if (ri == 0 || ci == 0 || ri == cellWidth - 1 || ci == cellHeight - 1) {
					cout << "*";
				}
			}
		}
		SetClr(7);
		Cell* temp = head;
		Cell* storeCurr = curr;
		int storeCurrRow = rCurr, storeCurrCol = cCurr;
		rCurr = 0;
		int ri = 0, ci = 0;
		while (!temp) {
			cCurr = 0;
			curr = temp;
			while (!curr) {
				printInCell(ri,ci);
				curr = curr->right;
				cCurr++;
				ci++;
			}
			ri++;
			
			temp = temp->down;
			rCurr++;
		}
	}
	void printCellBorder(int row, int col, int clr = 7, int adj =0) {

		SetClr(clr);
		for (int ri = 0; ri < 4 - adj; ri++) {
			for (int ci = 0; ci < 9; ci++) {
				if (ri == 0 || ci == 0 || ri == 3 - adj || ci == 8) {
					gotoRowCol(row + ri, col + ci);
					cout << "*";
				}
				else {
					gotoRowCol(row*3 + ri, col*8 + ci);
					//cout << " ";
				}
			}
		}
	}

	void printCellButtons(int row, int col, string name) {
		for (int ri = 0; ri < 3; ri++) {
			for (int ci = 0; ci < 8; ci++) {
				gotoRowCol(row + ri, col + ci);
				if (ri == 0 || ci == 0 || ri == 2 || ci == 7) {
					SetClr(15);
					cout << "*";
				}
				else if(ri == 1 && ci == 3) {
					SetClr(5);
					cout << name;
				}
			}
		}
	}

	void printValues() {
		int ri = 1, ci = 2;
		Cell* mainTemp = head;
		while (head->down != nullptr) {
			Cell* temp2 = head;
			while (head->right != nullptr) {
				gotoRowCol(ri, ci);
				if (head->data.size() == 0) {
					cout << " ";
				}
				cout << head->data;
				head = head->right;
				ci += 8;
			}
			ci = 2;
			ri += 3;
			head = temp2;
			head = head->down;
		}
		head = mainTemp;
		
	}

	
	void printInCell() {
		SetClr(7);
		int r = 0.75*rCurr * cellHeight + cellHeight / 3;
		int c = cCurr * cellWidth + cellWidth / 3;
		gotoRowCol(r, c);
		cout << curr->data;
	}
	void printInCell(int row, int col) {
		SetClr(7);
		gotoRowCol(row, col);
		cout << curr->data;
	}

	void navigate(int ascii) {
	       if (ascii == 100) { //right
			if (curr->right == nullptr) {
				insertColRight();
				printGrid();
			}

			printCellBorder(rCurr * 3, cCurr * 8, 7); // highlight next red
			curr = curr->right;
			cCurr++;
			printCellBorder(rCurr * 3, cCurr * 8, 4); //highlight previous back to white


			}
		else if (ascii == 97) {//left
				/*if (curr->left == nullptr) {
					insertColRight();
					printGrid();
				}*/
				printCellBorder(rCurr * 3, cCurr * 8, 7);
				if (curr->left != nullptr) {
					curr = curr->left;
					cCurr--;
				}
				printCellBorder(rCurr * 3, cCurr * 8, 4);
				}
		else if (ascii == 119) {//up
					/*if (curr->right == nullptr) {
						insertColRight();
						printGrid();
					}*/
					printCellBorder(rCurr * 3, cCurr * 8, 7);
					if (curr->up != nullptr) {
						curr = curr->up;
						rCurr--;
					}
					printCellBorder(rCurr * 3, cCurr * 8, 4);
					}
		else if (ascii == 115) {//down
						if (curr->down == nullptr) {
							insertrowDown();
							printGrid();
						}
						printCellBorder(rCurr * 3, cCurr * 8, 7);
						curr = curr->down;
						rCurr++;
						printCellBorder(rCurr * 3, cCurr * 8, 4);
						}
	}
	void inputKeyboard() {
		while (1) {
			char key = _getch();
			int ascii = key;
			if (ascii == 27) // Escape
				break;
			if (ascii == 8) {
				curr->data.pop_back();
			}

			if (ascii == 113) {// for input buttons

				int rI = -1, cI = -1;
				
				getRowColbyLeftClick(rI, cI);
				if ((rI >= 2 && rI <= 5) && (cI >= 74 && cI <= 82)) {
					curr->data = std::to_string(Sum());
				}
				else if ((rI >= 7 && rI <= 10) && (cI >= 74 && cI <= 82)) {
					curr->data = std::to_string(Average());
				}
				else if ((rI >= 12 && rI <= 15) && (cI >= 74 && cI <= 82)) {
					curr->data = std::to_string(Count());
				}
				else if ((rI >= 17 && rI <= 20) && (cI >= 74 && cI <= 82)) {
					curr->data = std::to_string(Max());
				}
				else if ((rI >= 21 && rI <= 24) && (cI >= 74 && cI <= 82)) {
					curr->data = std::to_string(Min());
				}
				else if ((rI >= 25 && rI <= 28) && (cI >= 74 && cI <= 82)) {
					save();
				}
				else if ((rI >= 30 && rI <= 34) && (cI >= 74 && cI <= 82)) {
					load();
					printGrid();
					printInCell();
					printValues();
					curr = head;
				}
			
			}
			if (rangeS || rangeE) {
				printCellBorder(rs * 3, cs * 8, 5);
				printCellBorder(re * 3, ce * 8, 5);
			}
			rs =cs=re=ce= 0;
			if (ascii == 100 || ascii == 97 || ascii == 119 || ascii == 115) {
				navigate(ascii);
			}
			
			else if(ascii == 102){// f to perform functions
			    Functions();
			}
			else if (ascii == 114) {
				if (!rangeStart) {
					rangeStart = curr;
					rs = rCurr; cs = cCurr;
					printCellBorder(rCurr * 3, cCurr * 8, 5);
					rangeS = true;
				}
				if (!rangeEnd) {
					int asciii=0;
					while(asciii!= 114){
						key = _getch();
						asciii = key;
						navigate(asciii);
					}
					rangeEnd = curr;
					re = rCurr; ce = cCurr;
					printCellBorder(rCurr * 3, cCurr * 8, 5);
					rangeE = true;
					
				}
			}
			

			// Insertions --------------------//
			else if (ascii == 111) {
				insertColLeft();
				printGrid();
			}
			else if (ascii == 112) {
				insertColRight();
				printGrid();
			}

			else if (ascii == 107) {
				insertRowUp();
				printGrid();
			}
			else if (ascii == 108) {
				insertrowDown();
				printGrid();
			}

			// Deletions -----------------//
			else if (ascii == 110) {
				deleteCol();
				system("cls");
				printGrid();
			}
			else if (ascii == 109) {
				deleteRow();
				system("cls");
				printGrid();
			}
			
			else if (ascii == 99) {
				rangeStart = nullptr;
				rangeEnd = nullptr;
				if (!rangeStart) {
					rangeStart = curr;
					rs = rCurr; cs = cCurr;
					printCellBorder(rCurr * 3, cCurr * 8, 5);
					rangeS = true;
				}
				if (!rangeEnd) {
					int asciii = 0;
					while (asciii != 114) {
						key = _getch();
						asciii = key;
						navigate(asciii);
					}
					rangeEnd = curr;
					re = rCurr; ce = cCurr;
					printCellBorder(rCurr * 3, cCurr * 8, 5);
					rangeE = true;

				}
				Copy();
			}
			else if (ascii == 120) {
				rangeStart = nullptr;
				rangeEnd = nullptr;
				if (!rangeStart) {
					rangeStart = curr;
					rs = rCurr; cs = cCurr;
					printCellBorder(rCurr * 3, cCurr * 8, 5);
					rangeS = true;
				}
				if (!rangeEnd) {
					int asciii = 0;
					while (asciii != 114) {
						key = _getch();
						asciii = key;
						navigate(asciii);
					}
					rangeEnd = curr;
					re = rCurr; ce = cCurr;
					printCellBorder(rCurr * 3, cCurr * 8, 5);
					rangeE = true;

				}
				Cut();
				printGrid();
				printValues();
			}
			else if (ascii == 118) {
				rangeStart = nullptr;
				rangeEnd = nullptr;
				if (!rangeStart) {
					rangeStart = curr;
					rs = rCurr; cs = cCurr;
					printCellBorder(rCurr * 3, cCurr * 8, 5);
					rangeS = true;
				}
				Paste();
				printGrid();
				printValues();
			}
			else if (ascii != 113) {
				char ch = ascii;
				if (curr->data.size() <= 4) {
					curr->data += ch;
				}
				printInCell();
				printValues();
			}
			
			printButtons();
			/*int rrow, ccol;
			getRowColbyLeftClick(rrow, ccol);
			gotoRowCol(rrow, ccol);
			cout << char(-37);
			cout << "\n r : " << rrow << " : " << ccol;*/
		}	
	}
	void Functions() {
		int rpos, cpos;
		getRowColbyLeftClick(rpos, cpos);
		rpos /= cellHeight,cpos/=cellWidth;
		if (rpos == 1 && cpos == 10) {
			printCellBorder(1, 10, 4);
			selectRange();
			inputKeyboard();
			//curr->data = Sum() + '0';
			printInCell();
			printCellBorder(1, 10, 7);
		}
		else if (rpos == 3 && cpos == 10) {
			printCellBorder(3, 10, 4);
			//selectRange();
			inputKeyboard();
			//curr->data = Average() + '0';
			printInCell();
			printCellBorder(3, 10, 7);
		}
		else if (rpos == 5 && cpos == 10) {
			printCellBorder(5, 10, 4);
			//selectRange();
			inputKeyboard();
			//curr->data = Count() + '0';
			printInCell();
			printCellBorder(5, 10, 7);
		}
		else if (rpos == 7 && cpos == 10) {
			printCellBorder(7, 10, 4);
			//selectRange();
			inputKeyboard();
			//curr->data = Min() + '0';
			printInCell();
			printCellBorder(7, 10, 7);
		}
		else if (rpos == 9 && cpos == 10) {
			printCellBorder(9, 10, 4);
			//selectRange();
			inputKeyboard();
			//curr->data = Max() + '0';
			printInCell();
			printCellBorder(9, 10, 7);
		}
		
	}
	void printButtons() {
		printCellButtons(2, 74,"SUM");
		printCellButtons(7, 74,"AVG");
		printCellButtons(12, 74,"CNT");
		printCellButtons(17, 74,"MAX");
		printCellButtons(21, 74,"MIN");
		printCellButtons(25, 74, "SAVE");
		printCellButtons(30, 74, "LOAD");
	} 

	void selectRange() {
		int ascii = 0;
		int selected = 0;
		while (selected < 2) {
			char key = _getch();
			ascii = key;
			if (ascii == 100) {//right
				if (curr->right != nullptr) {
					printCellBorder(rCurr, cCurr, 7);
					curr = curr->right;
					cCurr++;
					printCellBorder(rCurr, cCurr, 4);
				}
			}
			else if (ascii == 97) {//left
				printCellBorder(rCurr, cCurr, 7);
				if (curr->left != nullptr) {
					curr = curr->left;
					cCurr--;
				}
				printCellBorder(rCurr, cCurr, 4);
			}
			else if (ascii == 119) {//up
				printCellBorder(rCurr, cCurr, 7);
				if (curr->up != nullptr) {
					curr = curr->up;
					rCurr--;
				}
				printCellBorder(rCurr, cCurr, 4);
			}
			else if (ascii == 115) {//down
				printCellBorder(rCurr, cCurr, 7);
				if (curr->down != nullptr) {
					curr = curr->down;
					rCurr++;
				}
				printCellBorder(rCurr, cCurr, 4);
			}
			else if (ascii == 32) {// functions , space bar

				if (selected == 0) {
					rangeStart = curr;
					rs = rCurr;
					cs = cCurr;
					//printInCell(12, 10);
				}
				if (selected == 1) {
					rangeEnd = curr;
					re = rCurr;
					ce = cCurr;
					//printInCell(14,10);
				}
				selected++;
			}
			else {
				char ch = ascii;
				curr->data = ch;
				printInCell(); 
			}
		}
	}
	int Sum() {
		int sum = 0;
		Cell* temp = rangeStart;
		int rowIter = ceil(re - rs);
		int colIter = ce - cs;
		for (int ri = 0; ri <= rowIter; ri++) {
			Cell* temp2 = temp;
			for (int ci = 0; ci <= colIter; ci++) {
				sum += std::stoi(temp->data);
				temp = temp->right;
				
			}
			temp = temp2->down;
		}
		return sum;
	}
	double Average() {
		int sum = 0;
		int numOfelements = 0;
		Cell* temp = rangeStart;
		int rowIter = re - rs;
		int colIter = ce - cs;
		for (int ri = 0; ri <= rowIter; ri++) {
			Cell* temp2 = temp;
			for (int ci = 0; ci <= colIter; ci++) {
				sum += std::stoi(temp->data);
				temp = temp->right;
				numOfelements++;
			}
			temp = temp2->down;
		}
		return calculateAverage(sum, numOfelements);
		
	}
	int Count() {
		int count = 0;
		Cell* temp = rangeStart;
		int rowIter = re - rs;
		int colIter = ce - cs;
		for (int ri = 0; ri <= rowIter; ri++) {
			Cell* temp2 = temp;
			for (int ci = 0; ci <= colIter; ci++) {
				temp = temp->right;
				count++;
			}
			temp = temp2->down;
		}
		return count;
	}
	int Min() {
		vector<int> values;
		Cell* temp = rangeStart;
		int rowIter = re - rs;
		int colIter = ce - cs;
		for (int ri = 0; ri <= rowIter; ri++) {
			Cell* temp2 = temp;
			for (int ci = 0; ci <= colIter; ci++) {
				values.push_back(std::stoi(temp->data));
				temp = temp->right;
			}
			temp = temp2->down;
		}
		sort(values.begin(), values.end());
		if (values.size() > 0)
		   return values[0];
	}
	int Max() {
		vector<int> values;
		Cell* temp = rangeStart;
		int rowIter = re - rs;
		int colIter = ce - cs;
		for (int ri = 0; ri <= rowIter; ri++) {
			Cell* temp2 = temp;
			for (int ci = 0; ci <= colIter; ci++) {
				values.push_back(std::stoi(temp->data));
				temp = temp->right;
			}
			temp = temp2->down;
		}
		sort(values.begin(), values.end());
		if (values.size() > 0)
			return values[values.size()-1];
	}
	void save() {
		ofstream wrt("BSCS22137_Text.txt");
		wrt << rSize << " " << cSize;
		wrt << endl;
		Cell* temp = head;
		while (head->down != nullptr) {
			Cell* temp2 = head;
			while (head->right != nullptr) {
				wrt << head->data << " ";
				head = head->right;
			}
			head = temp2;
			wrt << endl;
			head = head->down;
		}
		head = temp;
	}
	void load() {
		ifstream rdr("BSCS22137_Text.txt");
		int row, col;
		rdr >> row >> col;
		string dat;
		int diffX = row - rSize;
		int diffY = col - cSize;

		Cell* temp = curr;
		for (int ri = 0; ri <= diffX; ri++) {
			insertrowDown();
			curr = curr->down;
		}
		for (int ci = 0; ci <= diffY; ci++) {
			insertColRight();
			curr = curr->right;
		}
		curr = temp;
		for (int ri = 0; ri < row; ri++) {
			Cell* temp2 = curr;
			for (int ci = 0; ci < col; ci++) {
				rdr >> dat;
				curr->data = dat;
				curr = curr->right;
			}
			curr = temp2;
			curr = curr->down;
		}
	}
	
	void Copy() {
		store.clear();
		Cell* temp = rangeStart;
		int rowIter = re - rs;
		int colIter = ce - cs;
		for (int ri = 0; ri <= rowIter; ri++) {
			Cell* temp2 = temp;
			for (int ci = 0; ci <= colIter; ci++) {
				store.push_back(temp->data);
				temp = temp->right;
			}
			temp = temp2->down;
		}
	}
	void Cut() {
		store.clear();
		Cell* temp = rangeStart;
		int rowIter = re - rs;
		int colIter = ce - cs;
		for (int ri = 0; ri <= rowIter; ri++) {
			Cell* temp2 = temp;
			for (int ci = 0; ci <= colIter; ci++) {
				store.push_back(temp->data);
				temp->data.clear();
				temp = temp->right;
			}
			temp = temp2->down;
		}
	}
	void Paste() {
		Cell* temp = rangeStart;
		int itr = store.size()-1;
		std::reverse(store.begin(), store.end());
		while (store.empty() == false && rangeStart->down != nullptr) {
			rangeStart->data += store[itr];
			store.pop_back();
			itr--;
			rangeStart = rangeStart->down;
		}
		rangeStart = temp;
		store.clear();
	}
	
};


