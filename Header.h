#pragma once
#include <iostream>
#include <Windows.h>
using namespace std;


template <typename Anytype>
class Cell
{
public:
	Anytype data;
	Cell<Anytype>* Up;
	Cell<Anytype>* Down;
	Cell<Anytype>* Left;
	Cell<Anytype>* Right;
	bool ActiveCell;
	Anytype row;
	Anytype Col;
	string color;

	Cell()
	{
		data = "X";
		Up = nullptr;
		Down = nullptr;
		Left = nullptr;
		Right = nullptr;
		ActiveCell = false;
		color = "\33[37m";
	}
	Cell(Anytype data)
	{
		this->data = data;
		Up = nullptr;
		Down = nullptr;
		Left = nullptr;
		Right = nullptr;
		ActiveCell = true;
	}

};
template <typename Anytype>
class Excel
{
public:
	Anytype rows;
	Anytype cols;
	Cell<Anytype>* Head;
	Cell<Anytype>* SelectedCell;


	Excel(){
	
		rows = cols = 5;
		Head = new Cell<Anytype>();
		SelectedCell = Head;

		this->makeGrid();

	}


	void makeGrid() {
		Cell<Anytype>* temporary = Head; // 1st cell at Head

		for (Anytype i = 1; i < rows; i++) // making starting column
		{
			Cell<Anytype>* NewCell = new Cell<Anytype>();
			temporary->Down = NewCell;
			NewCell->Up = temporary;
			temporary = temporary->Down;
		}

		temporary = Head;

		for (Anytype j = 1; j < cols; j++) {   // making starting row
			Cell<Anytype>* NewCell = new Cell <Anytype>();
			temporary->Right = NewCell;
			NewCell->Left = temporary;
			temporary = temporary ->Right;
		}
		temporary = Head;


		Cell<Anytype>* Previous = Head;
		Cell <Anytype>* Current = Head->Right;
		for (Anytype i = 1; i < cols; i++) {    // Creating Remaing Grid other than first row and first column //
			Cell<Anytype> * Prev = Previous->Down;
			Cell <Anytype>* Curr = Current;
			for (Anytype j = 1; j < rows; j++) {
				Cell<Anytype>* NewCell = new Cell<Anytype>();
				Prev->Right = NewCell;
				NewCell->Left = Prev;
				Prev = Prev->Down;
				Curr->Down = NewCell;
				NewCell->Up = Curr;
				Curr = NewCell;
			}
			Previous = Current;
			Current = Current->Right;
		}
	}

	void AboveInsert() {
		Cell<Anytype>* RowsHead = SelectedCell;
		while (RowsHead->Left != nullptr) {
			RowsHead = RowsHead->Left;
		}
		if (RowsHead->Up) {
			Cell<Anytype>* PreviousHead = RowsHead->Up;
			Cell<Anytype>* Previous = PreviousHead->Right;
			Cell <Anytype>* Current = RowsHead->Right;

			Cell<Anytype>* NewHead = new Cell <Anytype>();
			PreviousHead->Down = NewHead;
			NewHead->Up = PreviousHead;
			NewHead->Down = RowsHead;
			RowsHead->Up = NewHead;

			Cell<Anytype>* NewRowKaPointer = NewHead;
			for (Anytype i = 1; i < cols; i++) {

				Cell<Anytype>* NewCell = new Cell<Anytype>();

				NewRowKaPointer->Right = NewCell;
				NewCell->Left = NewRowKaPointer;

				Previous->Down = NewCell;
				NewCell->Up = Previous;

				Current->Up = NewCell;
				NewCell->Down = Current;

				NewRowKaPointer = NewCell;
				Previous = Previous->Right;
				Current = Current->Right;

			}
		}
		else {	// Now Making First row if there's no row
			Cell <Anytype>* NewRowHead = new Cell<Anytype>();
			Cell <Anytype>* Temporary = Head->Right;
			Head->Up = NewRowHead;
			NewRowHead->Down = Head;

			Cell <Anytype>* Current = NewRowHead;

			for (Anytype i = 1; i < cols; i++) {

				Cell <Anytype>* NeWcElL = new Cell<Anytype>();

				NeWcElL->Left = Current;
				Current->right = NeWcElL;
				Temporary->Up = NeWcElL;
				NeWcElL->Down = Temporary;
			    Temporary = Temporary->Right;
				Current = NeWcElL;
			}

			Head = NewRowHead;
		}
		rows += 1;

	}

	void BelowInsertion() {
		Cell<Anytype>* RowKaHEad = SelectedCell;
		while (RowKaHEad->Left != nullptr){

			RowKaHEad = RowKaHEad->Left;
		}

		if (RowKaHEad->Down) {

			Cell<Anytype>* NextRowHead = RowKaHEad->Down;
			Cell<Anytype>* NextPtr = NextRowHead->Right;
			Cell<Anytype>* Currentptr = RowKaHEad->Right;

			Cell<Anytype>* NewRowHead = new Cell <Anytype>();

			NextRowHead->Up = NewRowHead;
			NewRowHead->Down = NextRowHead;
			NewRowHead->Up = RowKaHEad;
			RowKaHEad->Down = NewRowHead;

			Cell<Anytype>* NewRowPTR = NewRowHead;

			for (Anytype i = 1 ; i< cols ; i++)
			{
				Cell <Anytype>* NewCell = new Cell <Anytype>();

				NewRowPTR->Right = NewCell;
				NewCell->Left = NewRowPTR;

				NextPtr->Up = NewCell;
				NewCell->Down = NextPtr;

				Currentptr->Down = NewCell;
				NewCell->Up = Currentptr;

				NewRowPTR = NewCell;
				NextPtr = NextPtr->Right;
				Currentptr = Currentptr->Right;

			}

		}
		else {
			NewRowAtLast(RowKaHEad); // Making bottom Row
		}
		rows += 1;
	}

	void NewRowAtLast(Cell<Anytype> *RowHead) {
		Cell<Anytype>* NewHeadRow = new Cell<Anytype>();
		Cell<Anytype>* Temporary = RowHead->Right;

		RowHead->Down = NewHeadRow;
		NewHeadRow->Up = RowHead;

		Cell<Anytype>* Currentptr = NewHeadRow;

		for (Anytype i = 1; i < cols; i++) {
			Cell<Anytype>* NewCell = new Cell<Anytype>();

			NewCell->Left = Currentptr;
			Currentptr->Right = NewCell;

			Temporary->Down = NewCell;
			NewCell->Up = Temporary;

			Temporary = Temporary->Right;
			Currentptr = NewCell;
		}
	}

	void InsertionAtRight() {

		Cell<Anytype>* ColumnHead = SelectedCell; // Taking first cell from selected Column

		while (ColumnHead->Up != nullptr) {

			ColumnHead = ColumnHead->Up;
		}

		if (ColumnHead->Right) {	// If there exist a column to the right of the Selected column 

			Cell<Anytype>* NewColumnHead = new Cell<Anytype>();
			Cell<Anytype>* NextColumnHead = ColumnHead->Right;

			Cell<Anytype>* NextPTR = NextColumnHead->Down;
			Cell<Anytype>* CurrentPTR = ColumnHead->Down;

			NewColumnHead->Left = ColumnHead;
			ColumnHead->Right = NewColumnHead;

			NewColumnHead->Right = NextColumnHead;
			NextColumnHead->Left = NewColumnHead;

			for (Anytype i = 1; i < rows; ++i) {

				Cell<Anytype>* NewCell = new Cell <Anytype>();

				NewColumnHead->Down = NewCell;
				NewCell->Up = NewColumnHead;

				NextPTR->Left = NewCell;
				NewCell->Right = NextPTR;

				CurrentPTR->Right = NewCell;
				NewCell->Left = CurrentPTR;

				NewColumnHead = NewCell;
				NextPTR = NextPTR->Down;
				CurrentPTR = CurrentPTR->Down;
			}
		}
		else {

			NewColumnAtRight(ColumnHead);	// If this this is last column from the selected columns
		}

		cols += 1;

	}
	void NewColumnAtRight(Cell<Anytype>* ColumnHead) {

		Cell<Anytype> *NewColumnkaHead = new Cell<Anytype>();
		Cell<Anytype>* Temporary = ColumnHead->Down;

		ColumnHead->Right = NewColumnkaHead;
		NewColumnkaHead->Left = ColumnHead;

		Cell<Anytype>* CurrentPTR = NewColumnkaHead;

		for (Anytype i = 1; i < rows; ++i) {
			Cell <Anytype>* NewCell = new Cell<Anytype>();

			CurrentPTR->Down = NewCell;
			NewCell->Up = CurrentPTR;

			Temporary->Right = NewCell;
			NewCell->Left = Temporary;

			CurrentPTR = NewCell;
			Temporary = Temporary->Down;
		}
	}
	void LeftInsertion() {

		Cell<Anytype>* ColumnHead = SelectedCell;
		while (ColumnHead->Up != nullptr)
		{
			ColumnHead = ColumnHead->Up;
		}
		if (ColumnHead->Left) { // If there exist a column to the left of the selected column

			Cell<Anytype>* NewColumnHEad = new Cell<Anytype>();
			Cell<Anytype>* PreviousHEad = ColumnHead->Left;
			Cell<Anytype>* PreviousPTR = PreviousHEad->Down;
			Cell<Anytype>* CurrentPTR = ColumnHead->Down;

			NewColumnHEad->Right = ColumnHead;
			ColumnHead->Left = NewColumnHEad;
			NewColumnHEad->Left = PreviousHEad;
			PreviousHEad->Right = NewColumnHEad;

			for (Anytype i = 1; i < rows; ++i) {

				Cell<Anytype>* NewCell = new Cell<Anytype>();

				NewColumnHEad->Down = NewCell;
				NewCell->Up = NewColumnHEad;

				PreviousPTR->Right = NewCell;
				NewCell->Left = PreviousPTR;

				CurrentPTR->Left = NewCell;
				NewCell->Right = CurrentPTR;

				NewColumnHEad = NewCell;
				PreviousPTR = PreviousPTR->Down;
				CurrentPTR = CurrentPTR->Down;


			}
		}
		else {
			// If there is no column at Left of Selected Cell
			NewColumnAtLeft(ColumnHead);
		}
		cols += 1;
	}
	void NewColumnAtLeft(Cell<Anytype>* ColumnHead) {

		Cell<Anytype>* NewColumnHead = new Cell<Anytype>();
		Cell<Anytype>* Temporary = ColumnHead->Down;

		ColumnHead->Left = NewColumnHead;
		NewColumnHead->Right = ColumnHead;
		Head = NewColumnHead;

		for (Anytype i = 1; i < rows; ++i) {

			Cell<Anytype>* NewCell = new Cell<Anytype>();

			NewColumnHead->Down = NewCell;
			NewCell->Up = NewColumnHead;

			Temporary->Left = NewCell;
			NewCell->Right = Temporary;

			NewColumnHead = NewCell;
			Temporary = Temporary->Down;
		}
	}
	void CellInsertionRightShift() {

		// insertion of new column at right most
		Cell<Anytype>* Temporary = SelectedCell;

		// Reaching to the right most end of the first row's last cell
		while (Temporary->Right != nullptr) {

			Temporary = Temporary->Right;
		}

		while (Temporary->Up != nullptr) {

			Temporary = Temporary->Up;
		}
		// Adding temporary as the top rightmost cell of New Column
		NewColumnAtRight(Temporary);

		// Right Shift All the Cells

		Cell<Anytype>* CurrentPTR = SelectedCell;
		Temporary = CurrentPTR->Left;

		Cell<Anytype>* NewCell = new Cell<Anytype>();
		if (Temporary) {
			Temporary->Right = NewCell;
			NewCell->Left = Temporary;
		}

		NewCell->Right = CurrentPTR;
		CurrentPTR->Left = NewCell;

		Temporary = NewCell;

		// Checking if it is first row

		if (!CurrentPTR) {

			Cell<Anytype>* Bottom = CurrentPTR->Down;

			if (SelectedCell == Head) {

				Head = NewCell;

			}

			while (Bottom) {

				Temporary->Down = Bottom;
				Bottom->UP = Temporary;

				Temporary = Temporary->Right;
				Bottom = Bottom->Right;

			}
		}
		else if (!CurrentPTR->Down) {

			// Checking for the Last Row
			Cell<Anytype>* Top = CurrentPTR->Up;

			while(Top) {

				Temporary->Up = Top;
				Top->Down = Temporary;

				Temporary = Temporary->Right;
				Top = Top->Right;

			}
		}
		else {

			// Checking for mid row

			Cell<Anytype>* Top = CurrentPTR->Up;
			Cell<Anytype>* Bottom = CurrentPTR->Down;

			while (Top) {

				Temporary->Up = Top;
				Top->Down = Temporary;

				Bottom->Up = Temporary;
				Temporary->Down = Bottom;

				Temporary = Temporary->Right;
				Top = Top->Right;
				Bottom = Bottom->Right;

			}

		}
		Temporary->Left->Right = nullptr;
		cols += 1;

	}
	void InsertionbyDownShift() {

		//inserting new row at last
		Cell<Anytype>* Temporary = SelectedCell;

		//bottom most left cell

		while (Temporary->Down != nullptr) {

			Temporary = Temporary->Down;

		}
		while(Temporary->Left != nullptr) {

			Temporary = Temporary->Left;

		}

		// Passing temporary as the bottom left most cell to Make a new row at bottom most
		NewRowAtLast(Temporary);

		//Shifting all the cells down to the selected cell
		Cell<Anytype>* CurrentPTR = SelectedCell;
		Temporary = CurrentPTR->Up;

		Cell<Anytype>* NewCell = new Cell<Anytype>();

		if (Temporary) {

			Temporary->Down = NewCell;
			NewCell->Up = Temporary;

		}
		NewCell->Down = CurrentPTR;
		CurrentPTR->Up = NewCell;

		Temporary = NewCell;

		if (!CurrentPTR->Left) {

			//checking for the first column
			Cell<Anytype>* right = CurrentPTR->Right;

			if (SelectedCell == Head) {

				Head = NewCell;

			}
			while (right) {

				Temporary->Right = right;
				right->Left = Temporary;

				Temporary = Temporary->Down;
				right = right->Down;

			}

		}
		else if (!CurrentPTR->Right) {

			//Checking for last last column
			Cell<Anytype>* left = CurrentPTR->Left;

			while (left) {

				Temporary->Left = left;
				left->Right = Temporary;

				Temporary = Temporary->Down;
				left = left->Down;
			}

		}
		else {

			// if it is middle row
			Cell<Anytype>* left = CurrentPTR->Left;
			Cell<Anytype>* right = CurrentPTR->Right;

			while (right) {

				Temporary->Right = right;
				right->Left = Temporary;

				left->Right = Temporary;
				Temporary->Left = left;

				Temporary = Temporary->Down;
				right = right->Down;
				left = left->Down;
			}

		}
		Temporary->Up->Down = nullptr;
		rows += 1;

	}

	void DeletionByLeftShift() {

		Cell<Anytype>* CurrentPTR = SelectedCell;
		if (!CurrentPTR->Left) {

			if (CurrentPTR->Right) {
				// Moving the data of the cell to the right
				CurrentPTR->data = CurrentPTR->Right->data;

				//Removing right cell
				Cell<Anytype>* Temporary = CurrentPTR->Right;
				CurrentPTR->Right = CurrentPTR->Right->Right;
				delete Temporary;
			}
			else {
				// If there exist only a single cell than remove it's data....
				CurrentPTR->data = Anytype();
			}

		}
		else if (!CurrentPTR->Right) {
			// if it is thr last column than delete it's data and update it's left pointer...
			Cell<Anytype>* Temporary = CurrentPTR;
			CurrentPTR = CurrentPTR->Left;
			CurrentPTR->Right = nullptr;
			delete Temporary;
		}
		else
		{	//if middle column shift the data of cell to the right
			CurrentPTR->data = CurrentPTR->Right->data;

			//Deleting Right Cell
			Cell<Anytype>* Temporary = CurrentPTR->Right;
			CurrentPTR->Right = CurrentPTR->Right->Right;
			delete Temporary;
		}

	}
	void printing() {
		Cell<Anytype>* Temporary = Head;
		Cell<Anytype>* Traverser;

		for (Anytype i = 0; i <= cols; ++i) {

			cout << i << " ";
		}
		cout << endl;

		for (Anytype i = 0; i < rows; ++i) {
			cout << i + 1 << "|";
			Traverser = Temporary;

			for (Anytype j = 0; j < cols; ++j) {

				cout << Traverser->color << Traverser->data << " |";
				Traverser = Traverser->Right;
			}
			cout << endl;
			Temporary = Temporary->Down;
		}
	}
	void Manual()
	{
		cout << "Use arrow keys to navigate" << endl;
		cout << "Use space to enter data" << endl;
		cout << "Use escape to exit" << endl;
		cout << "Press A to insert row above selected cell" << endl;
		cout << "Press B to insert row below selected cell" << endl;
		cout << "Press R to insert column to the right of selected cell" << endl;
		cout << "Press L to insert column to the left of selected cell" << endl;
		cout << "Press CTRL to insert cells by right shift" << endl;
		cout << "Press Shift to insert cells by down shift" << endl;
		cout << "Press Alt to delete cells by left shift" << endl;
	}
};