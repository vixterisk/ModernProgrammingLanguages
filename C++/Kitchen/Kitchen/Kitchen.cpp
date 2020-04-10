#include <iostream>
#include <string>
using namespace std;

//Кухонная утварь
class Kitchenware {
protected:
    int inventory_number;
public:
    Kitchenware(int number) 
    {
        inventory_number = number;
    }
    virtual void Print() = 0;
    void Print_inventory_number()
    {
        cout << "\ninventory number: " << inventory_number << "; \n";
    };
    int Get_inventory_number()
    {
        return inventory_number;
    }
};
//Плита
class Stove : public Kitchenware {
protected:
    string color;
    void Print(bool is_slow_cooker)
    {
        cout << "color: " << color << "; \n";
    }
public: 
    Stove(string stove_color, int number) : Kitchenware(number)
    {
        color = stove_color;
    }
    void Print()
    {
        Kitchenware::Print_inventory_number();
        cout << "color: " << color << "; \n";
    }
};
//Кастрюля
class Pan : public Kitchenware {
    int size;
public:
    Pan(int pan_size, int number) : Kitchenware(number)
    {
        size = pan_size;
    }

    void Print()
    {
        Kitchenware::Print_inventory_number();
        cout << "size: " << size << "; \n";
    }
};
//Плита электрическая
class Electric_stove : public Stove {
    int power;
protected:
    void Print(bool is_slow_cooker)
    {
        Stove::Print(is_slow_cooker);
        cout << "power: " << power << "; \n";
    }
public:
    Electric_stove(int stove_power, string stove_color, int number) : Stove(stove_color, number)
    {
        power = stove_power;
    }
    void Print()
    {
        Stove::Print();
        cout << "power: " << power << "; \n";
    }
};
//Плита газовая
class Gas_stove : public Stove {
    int consumption;
public:
    Gas_stove(int stove_consumption, string stove_color, int number) : Stove(stove_color, number)
    {
        consumption = stove_consumption;
    }
    void Print()
    {
        Stove::Print();
        cout << "consumption: " << consumption << "; \n";
    }
};
//Мультиварка
class Slow_cooker : public Electric_stove, public Pan {
    bool is_pressure_cooker;
public:
    Slow_cooker(bool pressure_cooker, int power, string color, int size, int number) : Electric_stove(power, color, number), Pan(size, number)
    {
        is_pressure_cooker = pressure_cooker;
    }
    void Print()
    {
        Pan::Print();
        Electric_stove::Print(true);
        if (is_pressure_cooker)
        cout << "This slow cooker is pressure cooker; \n";
        else
        cout << "This slow cooker is not a pressure cooker; \n";
    }
};
struct Node {
    Kitchenware* key;
    string kitchenware_type;
    int height;
    Node *left, *right;
};
typedef Node* PNode;

int Height(PNode p)
{
    return p == NULL ? 0 : p->height;
}

void NewHeight(PNode &p)
{
    int hl = Height(p->left);
    int hr = Height(p->right);
    int max = hl > hr ? hl : hr;
    p->height = max + 1;
}

int Say_balance(PNode p)
{
    return Height(p->right) - Height(p->left);
}

void LL(PNode& p)
{
    PNode left_son = p->left;
    p->left = left_son->right;
    left_son->right = p;
    NewHeight(p); p = left_son; NewHeight(p);
}

void RR(PNode& p)
{
    PNode right_son = p->right;
    p->right = right_son->left;
    right_son->left = p;
    NewHeight(p); p = right_son; NewHeight(p);
}

void RL(PNode& p) 
{
    LL(p->right);
    RR(p);
}

void LR(PNode& p)
{
    RR(p->left);
    LL(p);
}
//Балансировка
void Balance(PNode& p) 
{
    if (Say_balance(p) == 2) {
        if (Say_balance(p->right) >= 0) RR(p);
        else RL(p);
    }
    else if (Say_balance(p) == -2) {
        if (Say_balance(p->left) <= 0) LL(p);
        else LR(p);
    }
}
//Добавить в дерево
void Add_to_tree(PNode& Tree, Kitchenware* content, string type)
{
    if (Tree == NULL)
    {
        Tree = new Node;
        Tree->key = content;
        Tree->kitchenware_type = type;
        Tree->left = NULL;
        Tree->right = NULL;
    }
    else if (content->Get_inventory_number() < Tree->key->Get_inventory_number()) Add_to_tree(Tree->left, content, type);
    else Add_to_tree(Tree->right, content, type);
    NewHeight(Tree);
    Balance(Tree);
}
//Есть ли утварь с таким номером в дереве
bool Is_in_tree(PNode& Tree, int number) 
{
    if (Tree == NULL) return false;
    if (number < Tree->key->Get_inventory_number()) Is_in_tree(Tree->left, number);
    else if (number > Tree->key->Get_inventory_number()) Is_in_tree(Tree->right, number);
    else {
        cout << "Object with such number already exists: ";
        Tree->key->Print();
        cout << "\n";
        return true;
    }
}
//Удалить узел
void Delete(PNode& r, PNode& q)
{
    if (r->right != NULL) Delete(r->right, q);
    else {
        delete(q->key);
        q->key = r->key;
        q = r;
        r = r->left;
    }
}
//Удаление элемента из дерева
void Delete_from_tree(PNode& Tree, int number)
{
    PNode q;
    if (Tree == NULL) cout << "No such object. \n";
    else if (number < Tree->key->Get_inventory_number()) Delete_from_tree(Tree->left, number);
    else if (number > Tree->key->Get_inventory_number()) Delete_from_tree(Tree->right, number);
    else {
        q = Tree;
        if (q->right == NULL)
            Tree = q->left;
        else if (q->left == NULL)
            Tree = q->right;
        else Delete(q->left, q);
        delete(q);
    }
    if (Tree != NULL)
    {
        NewHeight(Tree);
        Balance(Tree);
    }
}
//Прямой обход
void Print1(PNode Tree)
{
    if (Tree == NULL) return;
    cout << "Kitchenware type - "<< Tree->kitchenware_type << ". Information: ";
    Tree->key->Print();
    cout << "\n";
    Print1(Tree->left);
    Print1(Tree->right);
}
//Обратный обход
void Print2(PNode Tree)
{
    if (Tree == NULL) return;
    Print2(Tree->left);
    cout << "Kitchenware type - " << Tree->kitchenware_type << ". Information: ";
    Tree->key->Print();
    cout << "\n";
    Print2(Tree->right);
}
//Концевой обход
void Print3(PNode Tree)
{
    if (Tree == NULL) return;
    Print3(Tree->left);
    Print3(Tree->right);
    cout << "Kitchenware type - " << Tree->kitchenware_type << ". Information: ";
    Tree->key->Print();
    cout << "\n";
}

void DeleteTree(PNode Tree)
{
    if (Tree == NULL) return;
    DeleteTree(Tree->left);
    DeleteTree(Tree->right);
    delete(Tree->key);
    delete(Tree);
}

//Ввод добавляемой утвари
void Add_kitchenware_to_tree(PNode &Tree) 
{
    string kitchenware;
    bool isOk = true;
    do {
        cout << "Enter kitchenware type (stove/pan/electric_stove/gas_stove/slow_cooker): ";
        cin >> kitchenware;
        if (kitchenware == "stove") {
            int number;
            cout << "Enter inventory number: ";
            cin >> number;
            if (Is_in_tree(Tree, number)) { isOk = false; continue; }
            string color;
            cout << "Enter color: ";
            cin >> color;
            Stove *stove = new Stove(color, number);
            Add_to_tree(Tree, (Kitchenware*)stove, "stove");
            isOk = true;
        }
        else if (kitchenware == "pan") {
            int number;
            cout << "Enter inventory number: ";
            cin >> number;
            if (Is_in_tree(Tree, number)) { isOk = false; continue; }
            int size;
            cout << "Enter size: ";
            cin >> size;
            Pan *pan = new Pan(size, number);
            Add_to_tree(Tree, (Kitchenware*)pan, "pan");
            isOk = true;
        }
        else if (kitchenware == "electric_stove") {
            int number;
            cout << "Enter inventory number: ";
            cin >> number;
            if (Is_in_tree(Tree, number)) { isOk = false; continue; }
            int power;
            cout << "Enter power: ";
            cin >> power; 
            string color;
            cout << "Enter color: ";
            cin >> color;
            Electric_stove* electric_stove = new Electric_stove(power, color, number);
            Add_to_tree(Tree, (Kitchenware*)electric_stove, "electric stove");
            isOk = true;
        }
        else if (kitchenware == "gas_stove") {
            int number;
            cout << "Enter inventory number: ";
            cin >> number;
            if (Is_in_tree(Tree, number)) { isOk = false; continue; }
            int consumption;
            cout << "Enter consumption: ";
            cin >> consumption;
            string color;
            cout << "Enter color: ";
            cin >> color;
            Gas_stove *gas_stove = new Gas_stove(consumption, color, number);
            Add_to_tree(Tree, (Kitchenware*)gas_stove, "gas stove");
            isOk = true;
        }
        else if (kitchenware == "slow_cooker") {
            int number;
            cout << "Enter inventory number: ";
            cin >> number;
            if (Is_in_tree(Tree, number)) { isOk = false; continue; }
            string input;
            cout << "Enter yes if is pressure cooker: ";
            cin >> input;
            bool is_pressure_cooker;
            is_pressure_cooker = input == "yes" ? true : false;
            int power;
            cout << "Enter power: ";
            cin >> power;
            string color;
            cout << "Enter color: ";
            cin >> color; 
            int size;
            cout << "Enter size: ";
            cin >> size;
            Slow_cooker *slow_cooker = new Slow_cooker(is_pressure_cooker, power, color, size, number);
            Add_to_tree(Tree, (Electric_stove*)slow_cooker, "slow cooker");
            isOk = true;
        }
        else isOk = false;
    } while (!isOk);
}

int main()
{   
    PNode Tree = NULL;
    string input;
    do {
        cout << "What would you like to do? (add/show/delete/exit to end program): ";
        cin >> input;
        if (input == "show")
        {
            cout << "What traversal? (pre-order/in-order/post-order): ";
            cin >> input; 
            cout << "\n";
            if (input == "pre-order") Print1(Tree);
            else if (input == "in-order") Print2(Tree);
            else Print3(Tree);
        }
        else if (input == "add") Add_kitchenware_to_tree(Tree);
        else if (input == "delete") { 
            int number;
            cout << "Enter inventory number of kitchenware to delete: ";
            cin >> number;
            Delete_from_tree(Tree, number); 
        }
    } while (input!="exit");
    DeleteTree(Tree);
    system("pause");
}