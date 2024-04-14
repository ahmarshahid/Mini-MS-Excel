#include <iostream>
#include <windows.h>
#include <conio.h>
#include <string>
#include <string.h>
#include <vector>
using namespace std;

template <typename Anytype>
class Cell {
public:
    Anytype data;
    Cell<Anytype>* up;
    Cell<Anytype>* down;
    Cell<Anytype>* left;
    Cell<Anytype>* right;
    bool ActiveCell;
    int row;
    int col;
    string color;

    Cell() {
        up = nullptr;
        down = nullptr;
        left = nullptr;
        right = nullptr;
        ActiveCell = false;
        data = 0;
        color = "\33[37m";
    }

    Cell(Anytype data) {
        this->data = data;
        up = nullptr;
        down = nullptr;
        left = nullptr;
        right = nullptr;
        ActiveCell = true;
    }
};

template <typename T>
class Excel {
public:
    int rows;
    int cols;
    Cell<T>* head;
    Cell<T>* selected;

    Excel() {
        rows = 5;
        cols = 5;
        head = new Cell<T>();
        selected = head;

        this->makeGrid();
        selected = head->right;
    }

    void makeGrid() {

        // after head cell
        Cell<T>* temp = head;

        // making first column
        for (int i = 1; i < rows; i++) {
            Cell<T>* newCell = new Cell<T>();
            temp->down = newCell;
            newCell->up = temp;
            temp = temp->down;
        }

        temp = head;

        // making first row
        for (int i = 1; i < cols; i++) {
            Cell<T>* newCell = new Cell<T>();
            temp->right = newCell;
            newCell->left = temp;
            temp = temp->right;
        }

        temp = head;

        // now making the rest of the grid
        Cell<T>* prevPointer = head;
        Cell<T>* currPointer = head->right;

        for (int i = 1; i < cols; i++) {
            Cell<T>* prev = prevPointer->down;
            Cell<T>* curr = currPointer;
            for (int j = 1; j < rows; j++) {
                Cell<T>* newCell = new Cell<T>();
                prev->right = newCell;
                newCell->left = prev;
                prev = prev->down;
                curr->down = newCell;
                newCell->up = curr;
                curr = newCell;
            }
            prevPointer = currPointer;
            currPointer = currPointer->right;
        }
    }

    void color(int c)
    {
        HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
        SetConsoleTextAttribute(hConsole, c);
        if (c > 255)
        {
            c = 1;
        }
    }



    void insertAbove() {

        // head onto the first cell of the selected row
        Cell<T>* rowHead = selected;
        while (rowHead->left != nullptr) {
            rowHead = rowHead->left;
        }

        if (rowHead->up) {
            // if there is a row above the selected row

            Cell<T>* prevRowHead = rowHead->up;
            Cell<T>* prevPointer = prevRowHead->right;
            Cell<T>* currPointer = rowHead->right;

            Cell <T>* newRowHead = new Cell<T>();
            prevRowHead->down = newRowHead;
            newRowHead->up = prevRowHead;
            newRowHead->down = rowHead;
            rowHead->up = newRowHead;

            Cell<T>* newRowPointer = newRowHead;

            for (int i = 1; i < cols; i++) {

                Cell<T>* newCell = new Cell<T>();
                newRowPointer->right = newCell;
                newCell->left = newRowPointer;

                prevPointer->down = newCell;
                newCell->up = prevPointer;

                currPointer->up = newCell;
                newCell->down = currPointer;

                newRowPointer = newCell;
                prevPointer = prevPointer->right;
                currPointer = currPointer->right;

            }
        }
        else {
            // this is the first row

            Cell<T>* newRowHead = new Cell<T>();
            Cell<T>* temp = head->right;
            head->up = newRowHead;
            newRowHead->down = head;

            Cell<T>* curr = newRowHead;
            for (int i = 1; i < cols; i++) {
                Cell<T>* newCell = new Cell<T>();
                newCell->left = curr;
                curr->right = newCell;
                temp->up = newCell;
                newCell->down = temp;
                temp = temp->right;
                curr = newCell;
            }

            head = newRowHead;
        }
        rows++;
    }

    void insertBelow() {

        // head onto the first cell of the selected row
        Cell<T>* rowHead = selected;
        while (rowHead->left != nullptr) {
            rowHead = rowHead->left;
        }

        if (rowHead->down) {
            // if there is a row below the selected row

            Cell<T>* nextRowHead = rowHead->down;
            Cell<T>* nextPointer = nextRowHead->right;
            Cell<T>* currPointer = rowHead->right;

            Cell <T>* newRowHead = new Cell<T>();
            nextRowHead->up = newRowHead;
            newRowHead->down = nextRowHead;
            newRowHead->up = rowHead;
            rowHead->down = newRowHead;

            Cell<T>* newRowPointer = newRowHead;

            for (int i = 1; i < cols; i++) {

                Cell<T>* newCell = new Cell<T>();
                newRowPointer->right = newCell;
                newCell->left = newRowPointer;

                nextPointer->up = newCell;
                newCell->down = nextPointer;

                currPointer->down = newCell;
                newCell->up = currPointer;

                newRowPointer = newCell;
                nextPointer = nextPointer->right;
                currPointer = currPointer->right;

            }
        }
        else {
            // this is the last row
            makeNewRowAtBottom(rowHead);

        }
        rows++;
    }

    void makeNewRowAtBottom(Cell<T>* rowHead) {
        Cell<T>* newRowHead = new Cell<T>();
        Cell<T>* temp = rowHead->right;
        rowHead->down = newRowHead;
        newRowHead->up = rowHead;

        Cell<T>* curr = newRowHead;
        for (int i = 1; i < cols; i++) {
            Cell<T>* newCell = new Cell<T>();
            newCell->left = curr;
            curr->right = newCell;
            temp->down = newCell;
            newCell->up = temp;
            temp = temp->right;
            curr = newCell;
        }
    }

    void insertRight() {

        // get the first cell of the selected column
        Cell<T>* colHead = selected;
        while (colHead->up != nullptr) {
            colHead = colHead->up;
        }

        if (colHead->right) {
            // if there is a column to the right of the selected column

            Cell<T>* nextColHead = colHead->right;
            Cell<T>* nextPointer = nextColHead->down;
            Cell<T>* currPointer = colHead->down;
            Cell<T>* newColHead = new Cell<T>();
            newColHead->left = colHead;
            colHead->right = newColHead;
            newColHead->right = nextColHead;
            nextColHead->left = newColHead;

            for (int i = 1; i < rows; i++) {
                Cell<T>* newCell = new Cell<T>();
                newColHead->down = newCell;
                newCell->up = newColHead;
                nextPointer->left = newCell;
                newCell->right = nextPointer;
                currPointer->right = newCell;
                newCell->left = currPointer;

                newColHead = newCell;
                nextPointer = nextPointer->down;
                currPointer = currPointer->down;
            }
        }
        else {
            // if it is the last column
            makeNewColumnAtRight(colHead);
        }

        cols++;
    }

    void makeNewColumnAtRight(Cell<T>* colHead) {
        Cell<T>* newColHead = new Cell<T>();
        Cell<T>* temp = colHead->down;
        colHead->right = newColHead;
        newColHead->left = colHead;
        Cell<T>* curr = newColHead;

        for (int i = 1; i < rows; i++) {
            Cell<T>* newCell = new Cell<T>();
            curr->down = newCell;
            newCell->up = curr;
            temp->right = newCell;
            newCell->left = temp;

            curr = newCell;
            temp = temp->down;
        }
    }

    void insertLeft() {

        // get the first cell of the selected column
        Cell<T>* colHead = selected;
        while (colHead->up != nullptr) {
            colHead = colHead->up;
        }

        if (colHead->left) {
            // if there is a column to the left of the selected column

            Cell<T>* prevColHead = colHead->left;
            Cell<T>* prevPointer = prevColHead->down;
            Cell<T>* currPointer = colHead->down;
            Cell<T>* newColHead = new Cell<T>();
            newColHead->right = colHead;
            colHead->left = newColHead;
            newColHead->left = prevColHead;
            prevColHead->right = newColHead;

            for (int i = 1; i < rows; i++) {
                Cell<T>* newCell = new Cell<T>();
                newColHead->down = newCell;
                newCell->up = newColHead;
                prevPointer->right = newCell;
                newCell->left = prevPointer;
                currPointer->left = newCell;
                newCell->right = currPointer;

                newColHead = newCell;
                prevPointer = prevPointer->down;
                currPointer = currPointer->down;
            }
        }
        else {
            // if this is the first column
            makeNewColumnAtLeft(colHead);

        }
        cols++;
    }

    void makeNewColumnAtLeft(Cell<T>* colHead) {
        Cell<T>* newColHead = new Cell<T>();
        Cell<T>* temp = colHead->down;
        colHead->left = newColHead;
        newColHead->right = colHead;
        head = newColHead;

        for (int i = 1; i < rows; i++) {
            Cell<T>* newCell = new Cell<T>();
            newColHead->down = newCell;
            newCell->up = newColHead;
            temp->left = newCell;
            newCell->right = temp;

            newColHead = newCell;
            temp = temp->down;
        }
    }

    void insertCellByRightShift() {

        // insert a new column at the right most end first
        Cell<T>* temp = selected;

        // getting to the right most end top cell
        while (temp->right != nullptr) {
            temp = temp->right;
        }

        while (temp->up != nullptr) {
            temp = temp->up;
        }

        // now temp is the top right most cell
        // inserting a new column at the right most end
        makeNewColumnAtRight(temp);

        // from selected cell, right shift all the cells
        Cell<T>* current = selected;
        temp = current->left;

        Cell<T>* newCell = new Cell<T>();
        if (temp) {
            temp->right = newCell;
            newCell->left = temp;
        }
        newCell->right = current;
        current->left = newCell;

        temp = newCell;

        // if it is the first row
        if (!current->up) {
            Cell<T>* below = current->down;

            if (selected == head) {
                head = newCell;
            }

            while (below) {
                temp->down = below;
                below->up = temp;

                temp = temp->right;
                below = below->right;
            }

        }
        else if (!current->down) {
            // if it is the last row
            Cell<T>* above = current->up;

            while (above) {
                temp->up = above;
                above->down = temp;

                temp = temp->right;
                above = above->right;
            }

        }
        else {
            // if it is a middle row
            Cell<T>* above = current->up;
            Cell<T>* below = current->down;

            while (above) {
                temp->up = above;
                above->down = temp;
                below->up = temp;
                temp->down = below;

                temp = temp->right;
                above = above->right;
                below = below->right;
            }

        }
        temp->left->right = nullptr;
        cols++;
    }

    void insertCellByDownShift() {
        // insert a new row at the bottom most end first
        Cell<T>* temp = selected;

        // get to the bottom most end left cell
        while (temp->down != nullptr) {
            temp = temp->down;
        }

        while (temp->left != nullptr) {
            temp = temp->left;
        }

        // now temp is the bottom left most cell
        // inserting a new row at the bottom most end
        makeNewRowAtBottom(temp);

        // from selected cell, down shift all the cells
        Cell<T>* current = selected;
        temp = current->up;

        Cell<T>* newCell = new Cell<T>();
        if (temp) {
            temp->down = newCell;
            newCell->up = temp;
        }
        newCell->down = current;
        current->up = newCell;

        temp = newCell;

        if (!current->left) {
            // it is the first column
            Cell<T>* right = current->right;

            if (selected == head) {
                head = newCell;
            }

            while (right) {
                temp->right = right;
                right->left = temp;

                temp = temp->down;
                right = right->down;
            }
        }
        else if (!current->right) {
            // it is the last column
            Cell<T>* left = current->left;

            while (left) {
                temp->left = left;
                left->right = temp;

                temp = temp->down;
                left = left->down;
            }
        }
        else {
            // it is a middle column
            Cell<T>* left = current->left;
            Cell<T>* right = current->right;

            while (right) {
                temp->right = right;
                right->left = temp;
                left->right = temp;
                temp->left = left;

                temp = temp->down;
                right = right->down;
                left = left->down;
            }
        }

        temp->up->down = nullptr;
        rows++;
    }

    void deleteCellByLeftShift()
    {
        Cell<T>* temp = selected;
        Cell<T>* prev = nullptr;

        while (temp != nullptr)
        {
            if (temp->right != nullptr)
            {
                temp->data = temp->right->data;
            }
            else
            {
                temp->data = 0;
            }

            prev = temp;
            temp = temp->right;
        }
    }

    void deleteCellByUpShift()
    {
        Cell<T>* Temp = selected;
        Cell<T>* Prev = nullptr;

        while (Temp != nullptr) {
            if (Temp->down != nullptr) {
                Temp->data = Temp->down->data;
            }
            else {
                Temp->data = 0;
            }
            Prev = Temp;
            Temp = Temp->down;
        }
    }
    void deleteColumn()
    {
        Cell<T>* Temporary = selected;
        selected = selected->left;
        while (Temporary->up != nullptr)
        {
            Temporary = Temporary->up;
        }
        if (Temporary->left != nullptr && Temporary->right != nullptr)
        {
            while (Temporary != nullptr)
            {
                Temporary->right->left = Temporary->left;
                Temporary->left->right = Temporary->right;
                Temporary->left = Temporary->right = Temporary->up = nullptr;
                Temporary = Temporary->down;
            }
        }
        else if (Temporary->left != nullptr)
        {
            while (Temporary != nullptr)
            {
                Temporary->left->right = nullptr;
                Temporary->left = Temporary->right = Temporary->up = nullptr;
                Temporary = Temporary->down;
            }
        }
        else {
            head = head->right;
            while (Temporary != nullptr)
            {
                Temporary->right->left = nullptr;
                Temporary->left = Temporary->right = Temporary->up = nullptr;
                Temporary = Temporary->down;
            }
        }
    }

    void deleteRow()
    {
        Cell<T>* temp = selected;
        selected = selected->up;
        while (temp->left != nullptr)
        {
            temp = temp->left;
        }
        if (temp->up != nullptr && temp->down != nullptr)
        {
            while (temp != nullptr)
            {
                temp->up->down = temp->down;
                temp->down->up = temp->up;
                temp->up = temp->down = temp->left = nullptr;
                temp = temp->right;
            }
        }
        else if (temp->up != nullptr)
        {
            while (temp != nullptr)
            {
                temp->up->down = nullptr;
                temp->up = temp->down = temp->left = nullptr;
                temp = temp->right;
            }
        }
        else
        {
            this->rootNode = this->rootNode->down;
            while (temp != nullptr)
            {
                temp->down->up = nullptr;
                temp->up = temp->down = temp->left = nullptr;
                temp = temp->right;
            }
        }
    }

    void ClearRow()
    {
        Cell<T>* Temporary = selected;
        while (Temporary->left != nullptr)
        {
            Temporary = Temporary->left;
        }
        while (Temporary != nullptr)
        {
            Temporary->data = -1;
            Temporary = Temporary->right;
        }
    }

    void ClearColumn()
    {
        Cell<T>* Temp = selected;
        while (Temp->up != nullptr)
        {
            Temp = Temp->up;
        }
        while (Temp != nullptr)
        {
            Temp->data = -1;
            Temp = Temp->down;
        }
    }

    T Sum(Cell<T>* Starting, Cell<T>* Ending)
    {
        int sum = 0;
        Cell<T>* RowNumber = Starting;
        while (RowNumber->left != nullptr) {
            RowNumber = RowNumber->left;
        }
        while (Starting != Ending) {
            sum = sum + Starting->data;

            if (Starting->right != nullptr) {

                Starting = Starting->right;
            }
            else if (RowNumber->down != nullptr)
            {
                RowNumber = RowNumber->down;
                Starting = RowNumber;
            }
            else
            {
                sum = Sum(Ending, Starting);
                return sum;
            }
        }
        sum += Ending->data;
        return sum;
    }

    T GetAverage(Cell<T>* beginPoint, Cell<T>* endingPoint)
    {
        int total = 0;
        T Avg = 0;
        T Sum = 0;
        Cell<T>* RowNumber = beginPoint;
        while (RowNumber->left != nullptr)
        {
            RowNumber = RowNumber->left;
        }
        while (beginPoint != endingPoint)
        {
            total += 1;
            Sum += beginPoint->data;
            if (beginPoint->right != nullptr)
            {
                beginPoint = beginPoint->right;
            }
            else if (RowNumber->down != nullptr)
            {
                RowNumber = RowNumber->down;
                beginPoint = RowNumber;
            }
            else
            {
                Avg = GetAverage(endingPoint, beginPoint);
                return Avg;
            }
        }
        Sum += endingPoint->data;
        total += 1;
        return Sum / total;
    }

    int count(Cell<T>* Begin, Cell<T>* Ending)
    {
        int total = 0;
        Cell<T>* RowLocation = Begin;
        while (RowLocation->left != nullptr)
        {
            RowLocation = RowLocation->left;
        }
        while (Begin != Ending)
        {
            total++;
            if (Begin->right != nullptr)
            {
                Begin = Begin->right;
            }
            else if (RowLocation->down != nullptr)
            {
                RowLocation = RowLocation->down;
                Begin = RowLocation;
            }
            else
            {
                return total;
            }
        }
        total++;
        return total;
    }

    T Minimum(Cell<T>* Start, Cell<T>* endingCell)
    {
        T minimum = Start->data;
        Cell<T>* rowNumber = Start;
        while (rowNumber->left != nullptr)
        {
            rowNumber = rowNumber->left;
        }
        while (Start != endingCell)
        {
            if (Start->data < minimum)
            {
                minimum = Start->data;
            }
            if (Start->right != nullptr)
            {
                Start = Start->right;
            }
            else if (rowNumber->down != nullptr)
            {
                rowNumber = rowNumber->down;
                Start = rowNumber;
            }
            else
            {
                return minimum;
            }
        }
        if (endingCell->data < minimum)
        {
            minimum = endingCell->data;
        }
        return minimum;
    }

    T Maximum(Cell<T>* Starting, Cell<T>* Ending)
    {
        T maximum = Starting->data;
        Cell<T>* RowNumber = Starting;
        while (RowNumber->left != nullptr)
        {
            RowNumber = RowNumber->left;
        }
        while (Starting != Ending)
        {
            if (Starting->data > maximum)
            {
                maximum = Starting->data;
            }
            if (Starting->right != nullptr)
            {
                Starting = Starting->right;
            }
            else if (RowNumber->down != nullptr)
            {
                RowNumber = RowNumber->down;
                Starting = RowNumber;
            }
            else
            {
                return maximum;
            }
        }
        if (Ending->data > maximum)
        {
            maximum = Ending->data;
        }
        return maximum;
    }

   /* vector<T> Cut(Cell<T>* StartingCell, Cell<T>* EndingCell)
    {
        vector<T> v;
        Cell<T>* RowNumber = StartingCell;
        while (RowNumber->left != nullptr)
        {
            RowNumber = RowNumber->left;
        }
        while (StartingCell != EndingCell)
        {
            v.push_back(StartingCell->data);
            StartingCell->data = -1;
            if (StartingCell->right != nullptr)
            {
                StartingCell = StartingCell->right;
            }
            else if (RowNumber->down != nullptr)
            {
                RowNumber = RowNumber->down;
                StartingCell = RowNumber;
            }
            else
            {
                return v;
            }
        }
        v.push_back(EndingCell->data);
        EndingCell->data = -1;
        return array;
    }

    vector<int> Copy(Cell<T>* startingCell, Cell<T>* endingCell)
    {
        vector<int> array;
        Cell<T>* rowNumber = startingcell;
        while (rowNumber->left != nullptr)
        {
            rowNumber = rowNumber->left;
        }
        while (starting != endingCell)
        {
            array.push_back(starting->data);
            if (starting->right != nullptr)
            {
                starting = starting->right;
            }
            else if (rowNumber->down != nullptr)
            {
                rowNumber = rowNumber->down;
                starting = rowNumber;
            }
            else
            {
                return array;
            }
        }
        array.push_back(endingCell->data);
        return array;
    }

    void Paste(Cell<T>* startingCell, vector<T> Array)
    {
        Cell<T>* rowNumber = startingCell;
        while (rowNumber->left != nullptr)
        {
            rowNumber = rowNumber->left;
        }
        for (int i = 0; i < vec.size(); i++)
        {
            startingCell->data = Array[i];
            if (startingCell->right != nullptr)
            {
                startingCell = startingCell->right;
            }
            else if (rowNumber->down != nullptr)
            {
                rowNumber = rowNumber->down;
                startingCell = rowNumber;
            }
            else
            {
                Cell<T>* temp = this->currentCell;
                this->currentCell = startingCell;
                insertRowBelow();
                rowNumber = rowNumber->down;
                startingCell = rowNumber;
                this->currentCell = temp;
            }
        }
    }*/


    void print() {
        Cell<T>* temp = head;
        Cell<T>* rowTraverser;

        //string horizontal_bar = ;

        // Print column headers
        cout << "   ";
        for (int i = 1; i <= cols; i++) {
            cout << "   " << i << "  ";
        }
        cout << std::endl;
        // Print horizontal line
        std::cout << "  *";
        for (int i = 0; i < cols; i++) {
            std::cout << "------";
        }
        std::cout << std::endl;

        // Print grid content
        for (int i = 0; i < rows; i++) {
            std::cout << i + 1 << " |";
            rowTraverser = temp;
            for (int j = 0; j < cols; j++) {
                std::cout << " " << rowTraverser->color << rowTraverser->data << " |";
                rowTraverser = rowTraverser->right;
            }
            std::cout << std::endl;

            // Print horizontal line
            std::cout << "  *";
            for (int k = 0; k < cols; k++) {
                std::cout << "------";
            }
            std::cout << std::endl;

            temp = temp->down;
        }
    }
};

void Manual() {
    cout << "Use following keys for different Functions : " << endl;
    cout << endl;
    cout << "Arrows for Movement" << endl;
    cout << "Space for Input" << endl;
    cout << "Escape to exit" << endl;
    cout << "A to insert row above selected cell" << endl;
    cout << "B to insert row below selected cell" << endl;
    cout << "R to insert column to the right of selected cell" << endl;
    cout << "L to insert column to the left of selected cell" << endl;
    cout << "CTRL to insert cells by right shift" << endl;
    cout << "Shift to insert cells by down shift" << endl;
    cout << "D to delete cell by left shift" << endl;
    cout << "Q to delete cell by up shift" << endl;
    cout << "K to delete column" << endl;
    cout << "S to Get Range Sum" << endl;
    cout << "Y to Get Average" << endl;
    cout << "O to Get Count" << endl;
    cout << "M to Get Minimum Number" << endl;
    cout << "N to Get Maximum Number" << endl;
    cout << "C to Copy" << endl;
    cout << "X to Cut" << endl;
    cout << "P to Paste" << endl;
    cout << endl;
}

//template <typename T>
int main() {

    Manual();
    bool running = true, modify = false;
    Excel<int>* excel = new Excel<int>();
    excel->print();

    Cell<int>* temp1 = nullptr, * temp2 = nullptr;
   // vector<int> vector = excel->Copy(temp1, temp2);
    while (running) {
        if (GetAsyncKeyState(VK_UP) && excel->selected->up) {
            excel->selected->color = "\33[37m";
            excel->selected = excel->selected->up;
            excel->selected->color = "\33[33m";
            modify = true;
        }
        if (GetAsyncKeyState(VK_DOWN) && excel->selected->down) {
            excel->selected->color = "\33[37m";
            excel->selected = excel->selected->down;
            excel->selected->color = "\33[33m";
            modify = true;
        }
        if (GetAsyncKeyState(VK_LEFT) && excel->selected->left) {
            excel->selected->color = "\33[37m";
            excel->selected = excel->selected->left;
            excel->selected->color = "\33[33m";
            modify = true;
        }
        if (GetAsyncKeyState(VK_RIGHT) && excel->selected->right) {
            excel->selected->color = "\33[37m";
            excel->selected = excel->selected->right;
            excel->selected->color = "\33[33m";
            modify = true;
        }
        if (GetAsyncKeyState(VK_SPACE)) {
            cout << "Enter data: ";
            int data;
            cin >> data;
            excel->selected->data = data;
            cin.ignore();

            modify = true;
        }
        else if (GetAsyncKeyState('A')) {
            excel->insertAbove();
            modify = true;
        }
        else if (GetAsyncKeyState('B')) {
            excel->insertBelow();
            modify = true;
        }
        else if (GetAsyncKeyState('R')) {
            excel->insertRight();
            modify = true;
        }
        else if (GetAsyncKeyState('L')) {
            excel->insertLeft();
            modify = true;
        }
        else if (GetAsyncKeyState(VK_CONTROL)) {
            excel->insertCellByRightShift();
            modify = true;
        }
        else if (GetAsyncKeyState(VK_SHIFT)) {
            excel->insertCellByDownShift();
            modify = true;
        }
        else if (GetAsyncKeyState('D')) {
            excel->deleteCellByLeftShift();
            modify = true;
        }
        else if (GetAsyncKeyState('Q')) {
            excel->deleteCellByUpShift();
            modify = true;
        }
        else if (GetAsyncKeyState('K')) {
            excel->deleteColumn();
            modify = true;
        }
        else if (GetAsyncKeyState('S'))
        {

            if (temp1 == nullptr)
            {
                temp1 = excel->selected;
            }
            else if (temp2 == nullptr)
            {
                temp2 = excel->selected;
            }
            else {
                excel->selected->data = excel->Sum(temp1, temp2);
                temp1 = nullptr;
                temp2 = nullptr;
            }
        }
        else if (GetAsyncKeyState('Y'))
        {

            if (temp1 == nullptr)
            {
                temp1 = excel->selected;
            }
            else if (temp2 == nullptr)
            {
                temp2 = excel->selected;
            }
            else {
                excel->selected->data = excel->GetAverage(temp1, temp2);
                temp1 = nullptr;
                temp2 = nullptr;
            }
        }
        else if (GetAsyncKeyState('O'))
        {

            if (temp1 == nullptr)
            {
                temp1 = excel->selected;
            }
            else if (temp2 == nullptr)
            {
                temp2 = excel->selected;
            }
            else {
                excel->selected->data = excel->count(temp1, temp2);
                temp1 = nullptr;
                temp2 = nullptr;
            }
        }
        else if (GetAsyncKeyState('M'))
        {

            if (temp1 == nullptr)
            {
                temp1 = excel->selected;
            }
            else if (temp2 == nullptr)
            {
                temp2 = excel->selected;
            }
            else {
                excel->selected->data = excel->Maximum(temp1, temp2);
                temp1 = nullptr;
                temp2 = nullptr;
            }
        }
        else if (GetAsyncKeyState('N'))
        {

            if (temp1 == nullptr)
            {
                temp1 = excel->selected;
            }
            else if (temp2 == nullptr)
            {
                temp2 = excel->selected;
            }
            else {
                excel->selected->data = excel->Minimum(temp1, temp2);
                temp1 = nullptr;
                temp2 = nullptr;
            }
        }
        else if (GetAsyncKeyState('C'))
        {

            if (temp1 == nullptr)
            {
                temp1 = excel->selected;
            }
            else if (temp2 == nullptr)
            {
                temp2 = excel->selected;
            }
            else {
              //  excel->selected->data = excel->Copy(temp1, temp2);
                temp1 = nullptr;
                temp2 = nullptr;
            }
        }
        else if (GetAsyncKeyState('X'))
        {

            if (temp1 == nullptr)
            {
                temp1 = excel->selected;
            }
            else if (temp2 == nullptr)
            {
                temp2 = excel->selected;
            }
            else {
               // excel->selected->data = excel->Cut(temp1, temp2);
                temp1 = nullptr;
                temp2 = nullptr;
            }
        }
        else if (GetAsyncKeyState('P'))
        {
            //excel->Paste(temp1, vector);
        }

        if (GetAsyncKeyState(VK_ESCAPE)) {
            running = false;
        }

        if (modify) {
            system("cls");
            Manual();
            excel->print();
            modify = false;
        }

        Sleep(100);
    }
    return 0;
}