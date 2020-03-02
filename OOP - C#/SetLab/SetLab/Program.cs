using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SetLab
{
    // Абстрактный класс Множества
    public abstract class Set
    {
        protected int upperBound;
        
        public abstract void IncludeElementInSet(int element);

        public abstract void ExcludeElementFromSet(int element);

        public abstract bool IsInSet(int element);

        public bool FillSet(string set)
        {
            int[] intElements = new int[0];
            string pattern = @"(\,)";
            set = Regex.Replace(set, pattern, String.Empty);
            var strSet = set.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            intElements = new int[strSet.Length];
            for (int i = 0; i < intElements.Length; i++)
                try
                {
                    intElements[i] = Int32.Parse(strSet[i]);
                    if (intElements[i] < 0) return false;
                }
                catch
                {
                    return false;
                }
            FillSet(intElements);
            return true;
        }

        public void FillSet(int[] set)
        {
                foreach (var element in set)
                    IncludeElementInSet(element);
        }

        public void PrintSet()
        {
            for (int i = 1; i <= upperBound; i++)
            {
                if (IsInSet(i)) Console.Write(i + " ");
            }
        }
    }

    // Исключение для класса Множества при превышении верхней границы
    public class SetUpperBoundExceededException : Exception
    {
        public SetUpperBoundExceededException() { }

        public SetUpperBoundExceededException(string message) : base(message) { }

        public SetUpperBoundExceededException(string message, Exception inner) : base(message, inner) { }
    }

    // Множества на основе булевого массива
    public class SimpleSet : Set
    {
        bool[] set;

        public SimpleSet(int upperLimit)
        {
            set = new bool[upperLimit];
            upperBound = upperLimit;
        }

        void HandleElement(int element, bool isElementIn)
        {
            if (element > upperBound) throw new SetUpperBoundExceededException("Element was out of range. Must be natural number and less than the upper bound of the Set");
            var elementNumber = element - 1;
            set[elementNumber] = isElementIn;
        }

        public override void IncludeElementInSet(int element)
        {
            HandleElement(element, true);
        }

        public override void ExcludeElementFromSet(int element)
        {
            HandleElement(element, false);
        }

        public override bool IsInSet(int element)
        {
            if (element > upperBound) return false;
            var elementNumber = element - 1;
            return set[elementNumber];
        }

        public static SimpleSet operator + (SimpleSet set1, SimpleSet set2)
        {
            var limit = Math.Max(set1.set.Length, set2.set.Length);
            var resultSet = new SimpleSet(limit);
            for (int i = 1; i <= limit; i++)
                if (set1.IsInSet(i) || set2.IsInSet(i)) resultSet.IncludeElementInSet(i);
            return resultSet;
        }

        public static SimpleSet operator * (SimpleSet set1, SimpleSet set2)
        {
            var limit = Math.Min(set1.set.Length, set2.set.Length);
            var resultSet = new SimpleSet(limit);
            for (int i = 1; i <= limit; i++)
                if (set1.IsInSet(i) && set2.IsInSet(i)) resultSet.IncludeElementInSet(i);
            return resultSet;
        }
    }

    // Множества на основе битового массива
    public class BitSet : Set
    {
        int[] set;

        public BitSet(int upperLimit)
        {
            set = new int[upperLimit / 32 + 1];
            upperBound = upperLimit;
        }

        public override void IncludeElementInSet(int element)
        {
            if (element > upperBound) throw new SetUpperBoundExceededException("Element was out of range. Must be natural number and less than the upper bound of the Set");
            var arrayIndex = element / 32;
            var bitIndex = element % 32 - 1;
            set[arrayIndex] = set[arrayIndex] | (1 << bitIndex);
        }

        public override void ExcludeElementFromSet(int element)
        {
            if (element > upperBound) throw new SetUpperBoundExceededException("Element was out of range. Must be natural number and less than the upper bound of the Set");
            var arrayIndex = element / 32;
            var bitIndex = element % 32 - 1;
            set[arrayIndex] = set[arrayIndex] & ~(1 << bitIndex);
        }

        public override bool IsInSet(int element)
        {
            if (element > upperBound) return false;
            var arrayIndex = element / 32;
            var bitIndex = element % 32 - 1;
            return ((set[arrayIndex] & (1 << bitIndex)) != 0);
        }

        public static BitSet operator + (BitSet set1, BitSet set2)
        {
            var maxArrayIndex = Math.Max(set1.set.Length, set2.set.Length);
            var upperBound = Math.Max(set1.upperBound, set2.upperBound);
            var resultSet = new BitSet(upperBound);
            for (int i = 0; i < maxArrayIndex; i++)
            {
                if (set1.set.Length > maxArrayIndex) resultSet.set[i] = set2.set[i];
                else if (set2.set.Length > maxArrayIndex) resultSet.set[i] = set1.set[i];
                else resultSet.set[i] = set1.set[i] | set2.set[i];
            }
            return resultSet;
        }

        public static BitSet operator * (BitSet set1, BitSet set2)
        {
            var minArrayIndex = Math.Min(set1.set.Length, set2.set.Length);
            var upperBound = Math.Max(set1.upperBound, set2.upperBound);
            var resultSet = new BitSet(upperBound);
            resultSet.upperBound = Math.Max(set1.upperBound, set2.upperBound);
            for (int i = 0; i < minArrayIndex; i++)
                resultSet.set[i] = set1.set[i] & set2.set[i];
            return resultSet;
        }
    }

    // Мультимножества на основе целочисленного массива
    public class MultiSet : Set
    {
        int[] set;

        public MultiSet(int upperLimit)
        {
            set = new int[upperLimit];
            upperBound = upperLimit;
        }

        public override void IncludeElementInSet(int element)
        {
            if (element > upperBound) throw new SetUpperBoundExceededException("Element was out of range. Must be natural number and less than the upper bound of the Set");
            var elementNumber = element - 1;
            set[elementNumber]++;
        }

        public override void ExcludeElementFromSet(int element)
        {
            if (element > upperBound) throw new SetUpperBoundExceededException("Element was out of range. Must be natural number and less than the upper bound of the Set");
            var elementNumber = element - 1;
            if (set[elementNumber] > 0)
                set[elementNumber]--;
        }

        public override bool IsInSet(int element)
        {
            if (element > upperBound) return false;
            var elementNumber = element - 1;
            return set[elementNumber] > 0;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Set set; string input = ""; int limit = 0; int[] intArrayFromFile = new int[0];
            Console.WriteLine("Введите L/l для логического представления множества,\nB/b для битового представления множества, \nлибо любые другие символы для мультимножества: ");
            var type = Console.ReadLine();
            Console.WriteLine("Введите F/f для чтения из файла, либо любые другие символы для чтения из строки: ");
            var inputType = Console.ReadLine();
            if (inputType.ToLower() == "f")
            {
                Console.WriteLine("Введите директорию: ");
                var dir = Console.ReadLine();
                if (File.Exists(dir))
                {
                    var words = File.ReadAllLines(dir);
                    limit = Int32.Parse(words[0]);
                    intArrayFromFile = new int[words.Length];
                    for (int i = 1; i < words.Length; i++)
                        intArrayFromFile[i] = Int32.Parse(words[i]);
                }
                else {
                    Console.WriteLine("Ошибка чтения файла. Программа завершит работу.");
                    return;
                }
            }
            else do
                {
                    do
                        Console.WriteLine("Введите максимум множества: ");
                    while (!Int32.TryParse(Console.ReadLine(), out limit));
                    Console.WriteLine("Введите элементы множества: ");
                    input = Console.ReadLine();
                } while (limit <= 0);
            if (type.ToLower() == "l") set = new SimpleSet(limit);
            else if (type.ToLower() == "b") set = new BitSet(limit);
            else set = new MultiSet(limit);
            bool isReadOk = true;
            if (inputType.ToLower() == "f")
                set.FillSet(intArrayFromFile);
            else
                isReadOk = set.FillSet(input);
            if (!isReadOk)  {
                Console.WriteLine("Ошибка чтения. Программа завершит работу.");
                Console.ReadKey();
                return;
            }
            string userChoice;
            do
            {
                Console.WriteLine(@"Введите команду и через пробел - элемент. Команды: ""print"" для печати множества, ""include""  для добавления элемента, ""exclude"" для удаления эелемента, ""find"" для проверки наличия элемента во множестве, ""union"" для объединения с другим множеством, ""intersect"" для пересечения с другим множеством,""exit"" для выхода из программы: ");
                input = Console.ReadLine();
                userChoice = input.Split()[0];
                int number;
                if (input.Split().Length > 1)
                {
                    try
                    {
                        number = Int32.Parse(input.Split()[1]);
                    }
                    catch
                    {
                        Console.WriteLine("Команда не была распознана, попробуйте снова:  \n");
                        continue;
                    }
                    if (userChoice == "include")
                    {
                        try
                        {
                            set.IncludeElementInSet(number);
                        }
                        catch (SetUpperBoundExceededException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }
                    else if (userChoice == "exclude")
                    {
                        try
                        {
                            set.ExcludeElementFromSet(number);
                        }
                        catch (SetUpperBoundExceededException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }
                    else if (userChoice == "find")
                    {
                        try
                        {
                            if (set.IsInSet(number)) Console.WriteLine("Элемент присутствует во множестве \n");
                            else Console.WriteLine("Элемент отсутствует во множестве \n");
                        }
                        catch (SetUpperBoundExceededException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }
                }
                else if (userChoice == "union" || userChoice == "intersect")
                {
                    if (type.ToLower() != "l" && type.ToLower()!= "b")
                    {
                        Console.WriteLine("Мультимножества не участвуют в пересечении и объединении.");
                        continue;
                    }
                    Set secondSet;
                    do
                    {
                        do
                            Console.WriteLine("Введите максимум второго множества: ");
                        while (!Int32.TryParse(Console.ReadLine(), out limit));
                        Console.WriteLine("Введите второе множество: ");
                        input = Console.ReadLine();
                    } while (limit <= 0);
                    if (type == "L" || type == "l") secondSet = new SimpleSet(limit);
                    else if (type == "B" || type == "b") secondSet = new BitSet(limit);
                    else secondSet = new MultiSet(limit);
                    isReadOk = secondSet.FillSet(input);
                    if (!isReadOk)
                    {
                        Console.WriteLine("Ошибка чтения. Программа завершит работу.");
                        Console.ReadKey();
                        return;
                    }
                    if (userChoice == "union")
                    {
                        if (type.ToLower() == "l") set = (SimpleSet)set + (SimpleSet)secondSet;
                        else if (type.ToLower() == "b") set = (BitSet)set + (BitSet)secondSet;
                    }
                    else if (userChoice == "intersect")
                    {
                        if (type.ToLower() == "l") set = (SimpleSet)set * (SimpleSet)secondSet;
                        else if (type.ToLower() == "b") set = (BitSet)set * (BitSet)secondSet;
                    }
                }
                else if (userChoice == "print")
                {
                    set.PrintSet();
                    Console.WriteLine();
                }
                else Console.WriteLine("Команда не была распознана, попробуйте снова: \n");
            } while (userChoice != "exit");
        }
    }
}
