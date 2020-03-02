using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Students
{
    [Serializable]
    public class Student
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public string Faculty { get; set; }

        public Student()
        { }

        public Student(string name, string surname, string faculty)
        {
            Name = name;
            Surname = surname;
            Faculty = faculty;
        }
    }

    [Serializable]
    static class Students
    {
        public static List<Student> MainList { get; set; }
        public static List<Student> StudentList { get; set; }
        public static List<Student> FoundList { get; set; }
        public static List<int> FoundIndicesInMainList { get; set; }
        public static int currentStudent { get; set; }

        public static void ChangeToFoundList()
        {
            StudentList = FoundList;
        }

        public static void ChangeToMainList()
        {
            StudentList = MainList;
            FoundList = null;
            FoundIndicesInMainList = null;
        }

        public static void SearchForStudent(string fieldName, string value)
        {
            FoundList = new List<Student>();
            FoundIndicesInMainList = new List<int>();
            for (int i = 0; i < MainList.Count; i++)
            {
                if (fieldName.ToLower() == "имя" && MainList[i].Name.ToLower() == value.ToLower() ||
                    fieldName.ToLower() == "фамилия" && MainList[i].Surname.ToLower() == value.ToLower() ||
                    fieldName.ToLower() == "факультет" && MainList[i].Faculty.ToLower() == value.ToLower())
                    {
                        FoundList.Add(MainList[i]);
                        FoundIndicesInMainList.Add(i);
                    }
            }
        }

        public static void DeserializeStudentArray(string path)
        {
            var formatter = new XmlSerializer(typeof(List<Student>));
            using (var stream = new FileStream(path, FileMode.Open))
            {
                MainList = StudentList = (List<Student>)formatter.Deserialize(stream);
                currentStudent = 0;
            }
        }

        public static void SerializeStudentArray(string path)
        {
            var formatter = new XmlSerializer(typeof(List<Student>));
            using (var stream = new FileStream(path, FileMode.Create))
            {
                formatter.Serialize(stream, StudentList);
            }
        }
    }
}
