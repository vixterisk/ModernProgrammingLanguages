using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace Students
{
    public partial class StudentsForm : Form
    {
        public StudentsForm()
        {
            InitializeComponent();
            Students.StudentList = new List<Student>();
            Students.currentStudent = 0;
            ClearStudentForm();
        }

        private void ClearStudentForm()
        {
            deleteStudentMenu.Enabled = false;
            HandlePrev(false);
            HandleNext(false);
            HandleTextBoxes(false);
            UpdateStudentInformation(String.Empty, String.Empty, String.Empty);
        }

        private void UpdateStudentInformation(string name, string surname, string faculty)
        {
            nameTextBox.Text = name;
            surnameTextBox.Text = surname;
            facultyTextBox.Text = faculty;
        }

        private void SelectStudentWithIndex(int index)
        {

            UpdateStudentInformation(Students.StudentList[index].Name,
                                     Students.StudentList[index].Surname,
                                     Students.StudentList[index].Faculty);
        }

        private void HandlePrev(bool isEnabled)
        {
            prevView.Enabled = isEnabled;
            prevButton.Enabled = isEnabled;
        }

        private void HandleNext(bool isEnabled)
        {
            nextView.Enabled = isEnabled;
            nextButton.Enabled = isEnabled;
        }

        private void nextView_Click(object sender, EventArgs e)
        {
            nextButton_Click(sender, e);
        }

        private void prevView_Click(object sender, EventArgs e)
        {
            prevButton_Click(sender, e);
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            Students.currentStudent++;
            SelectStudentWithIndex(Students.currentStudent);
            HandlePrev(true);
            if (Students.currentStudent == Students.StudentList.Count - 1)
                HandleNext(false);
        }

        private void prevButton_Click(object sender, EventArgs e)
        {
            Students.currentStudent--;
            SelectStudentWithIndex(Students.currentStudent);
            HandleNext(true);
            if (Students.currentStudent == 0)
                HandlePrev(false);
        }

        private void openStudentListMenu_Click(object sender, EventArgs e)
        {
            HandlePrev(false);
            HandleNext(false);
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                HandleTextBoxes(true);
                try
                {
                    Students.DeserializeStudentArray(openFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
                if (Students.StudentList.Count > 0)
                {
                    nameTextBox.Text = Students.StudentList[Students.currentStudent].Name;
                    surnameTextBox.Text = Students.StudentList[Students.currentStudent].Surname;
                    facultyTextBox.Text = Students.StudentList[Students.currentStudent].Faculty;
                    deleteStudentMenu.Enabled = true;
                    if (Students.StudentList.Count > 1)
                        HandleNext(true);
                }
            }

        }

        private void saveStudentListMenu_Click(object sender, EventArgs e)
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Students.SerializeStudentArray(saveFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void newStudentListMenu_Click(object sender, EventArgs e)
        {
            Students.StudentList = new List<Student>();
            ClearStudentForm();
        }

        private void addStudentMenu_Click(object sender, EventArgs e)
        {
            HandleTextBoxes(true);
            var newStudent = new Student("", "", "");
            Students.ChangeToMainList();
            Students.StudentList.Add(newStudent);
            Students.currentStudent = Students.StudentList.Count - 1;
            HandleNext(false);
            SelectStudentWithIndex(Students.currentStudent);
            deleteStudentMenu.Enabled = true;
            if (Students.StudentList.Count > 1)
                HandlePrev(true);
        }

        private void nameTextBox_TextChanged(object sender, EventArgs e)
        {
            if (Students.StudentList != null && Students.StudentList.Count > 0)
                Students.StudentList[Students.currentStudent].Name = nameTextBox.Text;
        }

        private void surnameTextBox_TextChanged(object sender, EventArgs e)
        {
            if (Students.StudentList != null && Students.StudentList.Count > 0)
                Students.StudentList[Students.currentStudent].Surname = surnameTextBox.Text;
        }

        private void facultyTextBox_TextChanged(object sender, EventArgs e)
        {
            if (Students.StudentList != null && Students.StudentList.Count > 0)
                Students.StudentList[Students.currentStudent].Faculty = facultyTextBox.Text;
        }

        private void deleteStudentMenu_Click(object sender, EventArgs e)
        {
            if (Students.FoundIndicesInMainList != null)
            {
                Students.MainList.RemoveAt(Students.FoundIndicesInMainList[Students.currentStudent]);
                Students.FoundIndicesInMainList.RemoveAt(Students.currentStudent);
            }
            Students.StudentList.RemoveAt(Students.currentStudent);
            if (Students.StudentList.Count == 0)
            {
                ClearStudentForm();
                return;
            }
            if (Students.currentStudent == Students.StudentList.Count)
                Students.currentStudent--;
            SelectStudentWithIndex(Students.currentStudent);
            if (Students.currentStudent == Students.StudentList.Count - 1)
                HandleNext(false);
            if (Students.currentStudent == 0)
                HandlePrev(false);
        }

        private void HandleTextBoxes(bool isEnabled)
        {
            nameTextBox.Enabled = isEnabled;
            surnameTextBox.Enabled = isEnabled;
            facultyTextBox.Enabled = isEnabled;
        }

        private void SelectFirstInList()
        {
            HandleTextBoxes(true);
            HandlePrev(false);
            Students.currentStudent = 0;
            SelectStudentWithIndex(Students.currentStudent);
            HandleNext(Students.StudentList.Count != 1);
        }

        private void ComboBoxAndTextBoxChanged()
        {
            HandlePrev(false);
            HandleNext(false);
            Students.SearchForStudent(searchComboBox.Text, searchTextBox.Text);
            if (Students.FoundList.Count > 0)
            {
                Students.ChangeToFoundList();
                SelectFirstInList();
            }
            else
                HandleTextBoxes(false);
        }

        private void searchComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (searchTextBox.Text != "")
                ComboBoxAndTextBoxChanged();
        }

        private void searchTextBox_TextChanged(object sender, EventArgs e)
        {
            if (searchComboBox.Text != "")
            {
                ComboBoxAndTextBoxChanged();
                if (searchTextBox.Text == "")
                {
                    Students.ChangeToMainList();
                    SelectFirstInList();
                }
            }
        }
    }
}
