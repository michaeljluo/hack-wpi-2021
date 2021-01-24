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



namespace Goat_Therapy_WPI_Hackathon
{
    public partial class Home_Page : Form
    {
        int survey_count = 1;

        public Home_Page()
        {
            InitializeComponent();
        }
 
        private void Home_Page_Load(object sender, EventArgs e)
        {
            load_club_list();
            watch();
            System.Diagnostics.Debug.WriteLine("Starting Program");
            this.check_1.Visible = true;
            this.check_2.Visible = true;
            this.check_3.Visible = true;
            this.check_4.Visible = true;
            this.check_5.Visible = true;
            this.check_6.Visible = false;

            this.check_1.Text = "What club is best for you?";
            this.check_2.Text = "Feeling Unhealthy?";
            this.check_3.Text = "What to do at WPI";
            this.check_4.Text = "I want to learn about Online Communities at WPI.";
            this.check_5.Text = "I want to try something new!";
            this.button_start.Text = "Take Survey";


        }
        string path = @"C:\Users\Zach\Pictures\Hackathon";
        private void watch() // Function to watch the folder
        {
            FileSystemWatcher watcher = new FileSystemWatcher(); // Set up the instance of a file watcher
            watcher.Path = path; // Select the desired path 
            watcher.NotifyFilter = NotifyFilters.LastWrite; // Define what the program should look for
            watcher.Filter = "*.*"; // Filter the results, since there is nothing in the folder this is broad 
            watcher.Changed += new FileSystemEventHandler(OnChanged); // Add an instance to what would happen if the desired scenario occurs
            watcher.EnableRaisingEvents = true; // Allow the program to take action when the event occurs
        }
        string resource1 = "";
        string resource2 = "";
        string resource3 = "";
        string resource4 = "";
        public void OnChanged(object source, FileSystemEventArgs e) // Reaction when the file is changed
        {
            System.Diagnostics.Debug.WriteLine("Change");
            /* 
             * Since the watcher functions are on a different thread then the excel doc, the functions cannot be called directly
             * This function in tandem with an if statement allows the program to switch between watching and reading the document
             */
            int file_count = Directory.GetFiles(path, "*", SearchOption.TopDirectoryOnly).Length;
            int y = 0;
         
            System.Diagnostics.Debug.WriteLine(file_count);
            for (int x = 0; x < file_count; x++)
            {
                if (e.FullPath.Contains("Survey_Result") && y < 4)
                {
                    switch (y)
                    {
                        case 0:
                            resource1 = e.FullPath;
                            y++;
                            break;
                        case 1:
                            resource2 = e.FullPath;
                            y++;
                            break;
                        case 2:
                            resource3 = e.FullPath;
                            y++;
                            break;
                        case 3:
                            resource4 = e.FullPath;
                            y++;
                            break;
                    }
                }
                else if (y >= 4) { return; }

                update(resource1, resource2, resource3, resource4);
            }
            
        }
        void update(string inp1, string inp2, string inp3, string inp4)
        {
            System.Diagnostics.Debug.WriteLine(resource1);
            this.linkLabel1.Text = inp1;
            this.linkLabel2.Text = inp2;
            this.linkLabel3.Text = inp3;
            this.linkLabel4.Text = inp4;
        }
        private void button_start_Click(object sender, EventArgs e)
        {
            survey_StateMachine();

        }

        private void check_1_CheckedChanged(object sender, EventArgs e)
        {
            if(check_1.Checked)
            {
                check_2.Checked = false;
                check_3.Checked = false;
                check_4.Checked = false;
                check_5.Checked = false;
                check_6.Checked = false;
            }
        }

        private void check_2_CheckedChanged(object sender, EventArgs e)
        {
            if (check_2.Checked)
            {
                check_1.Checked = false;
                check_3.Checked = false;
                check_4.Checked = false;
                check_5.Checked = false;
                check_6.Checked = false;
            }
        }

        private void check_3_CheckedChanged(object sender, EventArgs e)
        {
            if (check_3.Checked)
            {
                check_2.Checked = false;
                check_1.Checked = false;
                check_4.Checked = false;
                check_5.Checked = false;
                check_6.Checked = false;
            }
        }

        private void check_4_CheckedChanged(object sender, EventArgs e)
        {
            if (check_4.Checked)
            {
                check_2.Checked = false;
                check_3.Checked = false;
                check_1.Checked = false;
                check_5.Checked = false;
                check_6.Checked = false;
            }
        }

        private void check_5_CheckedChanged(object sender, EventArgs e)
        {
            if (check_5.Checked)
            {
                check_2.Checked = false;
                check_3.Checked = false;
                check_4.Checked = false;
                check_1.Checked = false;
                check_6.Checked = false;
            }
        }

        private void check_6_CheckedChanged(object sender, EventArgs e)
        {
            if (check_6.Checked)
            {
                check_2.Checked = false;
                check_3.Checked = false;
                check_4.Checked = false;
                check_5.Checked = false;
                check_1.Checked = false;
            }
        }

       void survey_StateMachine ()
        {
            switch (survey_count) {
                case 0:
                    this.check_1.Visible = true;
                    this.check_2.Visible = true;
                    this.check_3.Visible = true;
                    this.check_4.Visible = true;
                    this.check_5.Visible = true;
                    this.check_6.Visible = false;

                    this.check_1.Text = "What club is best for you?";
                    this.check_2.Text = "Feeling Unhealthy?";
                    this.check_3.Text = "What to do at WPI";
                    this.check_4.Text = "I want to learn about Online Communities at WPI.";
                    this.check_5.Text = "I want to try something new!";
                    this.button_start.Text = "Take Survey";
                    survey_count++;
                    break;
                case 1:
                    
                    if (check_1.Checked) // What club
                    {
                        survey_count = 2;
                        this.label_survey.Text = "What type of club interest you?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Clubs based on Social Interaction";
                        this.check_2.Text = "Clubs based on Academics";
                        this.check_3.Text = "Physical Club and Sports";
                    }
                    else if(check_2.Checked) // Feeling Unhealthy
                    {
                        survey_count = 7;
                        this.button_start.Text = "Next";
                        this.label_survey.Text = "What struggles are you facing?";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = false;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Mental Health related";
                        this.check_2.Text = "Physical Health related";
                    }
                    else if (check_3.Checked) // What to do at WPI
                    {
                        survey_count = 11;
                        this.label_survey.Text = "What type of activities interest you?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Do you want Off-Campus Activities?";
                        this.check_2.Text = "Do you want On-Campus Activities?";
                        this.check_3.Text = "Do you want an at-home/indoor activity?";
                    }
                    else if (check_4.Checked) // Online Comm
                    {
                        survey_count = 18;
                        this.label_survey.Text = "What type of community interest you?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = true;
                        this.check_5.Visible = true;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Class associated community";
                        this.check_2.Text = "Major related community";
                        this.check_3.Text = "Housing/Dorm community";
                        this.check_4.Text = "Club related community";
                        this.check_5.Text = "Video Game Related Community";
                    }
                    else if(check_5.Checked) // Try something new
                    {
                        survey_count = 11;
                        this.label_survey.Text = "What type of activities interest you?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Do you want Off-Campus Activities?";
                        this.check_2.Text = "Do you want On-Campus Activities?";
                        this.check_3.Text = "Do you want an at-home/indoor activity?";
                        this.check_4.Text = "Are there any clubs that interest you?";
                    }
                    else // Nothing entered the program does not change states
                    {
                        survey_count = 0;
                    }
                    break;
                case 2:

                    if (check_1.Checked) // Social
                    {
                        survey_count = 3;
                        this.label_survey.Text = "What type of social club interest you?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Greek Life";
                        this.check_2.Text = "General Interest Clubs";
                        this.check_3.Text = "Cultural/Religious Clubs";
                    }
                    else if (check_2.Checked) // Academic
                    {
                        survey_count = 5;
                        this.button_start.Text = "Next";
                        this.label_survey.Text = "What type of academic club interest you?";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = false;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Honor Societies";
                        this.check_2.Text = "Major Related";
                    }
                    else if (check_3.Checked) // Sport
                    {
                        survey_count = 6;
                        this.label_survey.Text = "What type of sport activities interest you?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Do you want Varsity Sports?";
                        this.check_2.Text = "Do you want Intramural Sports?";
                        this.check_3.Text = "Do you want Club Sports?";
                    }
                    break;
                case 3:

                    if (check_1.Checked) // Greek
                    {
                        survey_count = 4;
                        this.label_survey.Text = "What type of Greek Life are you interested in?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = false;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Fraternities (Brothers)";
                        this.check_2.Text = "Sororities (Sisters)";
                    }
                    else if (check_2.Checked) // general
                    {
                        resource_search('g');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_3.Checked) // cultural
                    {
                        resource_search('c');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 4:

                    if (check_1.Checked) // frat
                    {
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        resource_search('f');
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // girl
                    {
                        resource_search('s');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 5:

                    if (check_1.Checked) // honor
                    {
                        resource_search('h');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // major
                    {
                        resource_search('m');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 6:

                    if (check_1.Checked) // varsity
                    {
                        resource_search('v');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // intramural
                    {
                        resource_search('z');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if(check_3.Checked) // club
                    {
                        resource_search('S');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 7:

                    if (check_1.Checked) // Mental
                    {
                        survey_count = 8;
                        this.label_survey.Text = "What best describes you?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "I am stressed";
                        this.check_2.Text = "I have race related concerns";
                        this.check_3.Text = "I suffer from Depression/Anxiety";
                    }
                    else if (check_2.Checked) // Physical
                    {
                        survey_count = 10;
                        this.button_start.Text = "Next";
                        this.label_survey.Text = "Where can we help?";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = false;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Fitness";
                        this.check_2.Text = "Nutrition";
                    }
                    break;
                case 8:

                    if (check_1.Checked) // stress
                    {
                        survey_count = 9;
                        this.button_start.Text = "Next";
                        this.label_survey.Text = "What's the stressor?";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = false;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Academics";
                        this.check_2.Text = "Homesickness";
                    }
                    else if (check_2.Checked) // poc
                    {
                        resource_search('p');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_3.Checked) // dep
                    {
                        resource_search('d');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 9:

                    if (check_1.Checked) // academic
                    {
                        resource_search('a');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // homesick
                    {
                        resource_search('H');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 10:

                    if (check_1.Checked) // fit
                    {
                        resource_search('F');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // nutrition
                    {
                        resource_search('n');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 11:

                    if (check_1.Checked) // on-campus
                    {
                        survey_count = 12;
                        this.label_survey.Text = "What building/area do you want?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = true;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Campus Center";
                        this.check_2.Text = "Foisie";
                        this.check_3.Text = "Clubs";
                        this.check_4.Text = "Workspaces";
                    }
                    else if (check_2.Checked) // off-campus
                    {
                        survey_count = 13;
                        this.button_start.Text = "Next";
                        this.label_survey.Text = "What type of activity?";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = true;
                        this.check_5.Visible = false; 
                        this.check_6.Visible = false;
                        this.check_1.Text = "Something fun!";
                        this.check_2.Text = "Shopping";
                        this.check_3.Text = "Food";
                        this.check_4.Text = "Events";
                    }
                    else if (check_3.Checked) // at-home
                    {
                        survey_count = 17;
                        this.label_survey.Text = "What type of in-home activity interest you?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Something Online";
                        this.check_2.Text = "Arts and Crafts?";
                        this.check_3.Text = "Games";
                    }
                    break;
                case 12:

                    if (check_1.Checked) // cc
                    {
                        resource_search('C');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // foisie
                    {
                        resource_search('o');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_3.Checked) // clubs
                    {
                        survey_count = 2;
                        this.label_survey.Text = "What type of club interest you?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = true;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Clubs based on Social Interaction";
                        this.check_2.Text = "Clubs based on Academics";
                        this.check_3.Text = "Physical Club and Sports";
                    }
                    else if (check_4.Checked) // workspaces
                    {
                        resource_search('w');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 13:

                    if (check_1.Checked) // fun
                    {
                        survey_count = 14;
                        this.label_survey.Text = "What type of fun activity?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = false;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Other";
                        this.check_2.Text = "Physical Activities/Outdoors";
                    }
                    else if (check_2.Checked) // shopping
                    {
                        resource_search('G');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_3.Checked) // food
                    {
                        survey_count = 15;
                        this.label_survey.Text = "Cooking or Restaurant?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = false;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Cooking";
                        this.check_2.Text = "Restaurant";
                    }
                    else if (check_4.Checked) // events
                    {
                         survey_count = 16;
                        this.label_survey.Text = "What type of event?";
                        this.button_start.Text = "Next";
                        this.check_1.Visible = true;
                        this.check_2.Visible = true;
                        this.check_3.Visible = false;
                        this.check_4.Visible = false;
                        this.check_5.Visible = false;
                        this.check_6.Visible = false;
                        this.check_1.Text = "Festivals";
                        this.check_2.Text = "Concerts";
                    }
                    break;
                case 14:

                    if (check_1.Checked) // other
                    {
                        resource_search('O');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // physical
                    {
                        resource_search('T');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 15:

                    if (check_1.Checked) // cooking
                    {
                        resource_search('k');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // restaurant
                    {
                        resource_search('R');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 16:

                    if (check_1.Checked) // festival
                    {
                        resource_search('i');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // concerts
                    {
                        resource_search('e');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 17:

                    if (check_1.Checked) // online
                    {
                        resource_search('l');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // arts
                    {
                        resource_search('r');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_3.Checked) // game
                    {
                        resource_search('M');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
                case 18:

                    if (check_1.Checked) // class
                    {
                        resource_search('L');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_2.Checked) // major
                    {
                        resource_search('j');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_3.Checked) // house
                    {
                        resource_search('y');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_4.Checked) // clubs
                    {
                        resource_search('Y');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    else if (check_5.Checked) // game
                    {
                        resource_search('Z');
                        survey_count = 0;
                        this.button_start.Text = "Home";
                        //Case Switch to show resources
                    }
                    break;
            }
        }
        string excelfile = @"C:\Users\Zach\Downloads\Calc 3 Downloads\Resources.xlsx"; // The directory for the excel doc
        string sheetname = ""; // The name of the sheet
        int placehold;
        string necessary_group;
        void resource_search (char input)
        {
            Cursor.Current = Cursors.WaitCursor;
            Excel xlapp = new Excel();           
            switch (input) {
                case 'g': //general interest club
                    sheetname = "Clubs";
                    placehold = 13;
                    necessary_group = "general interest";
                    break;
                case 'c': // Cultural club
                    sheetname = "Clubs";
                    placehold = 13;
                    necessary_group = "Cultural club";
                    break;
                case 'f': //fraternity
                    sheetname = "Fraternities";
                    placehold = 9;
                    necessary_group = "CHECK";
                    break;
                case 'h': // Honor Societies
                    break;
                case 's': //Sorority
                    sheetname = "Sororities";
                    placehold = 10;
                    necessary_group = "CHECK";
                    break;
                case 'm': // major specific club
                    sheetname = "Clubs";
                    placehold = 13;
                    necessary_group = "Academic interest";
                    break;
                case 'v': //varsity
                    sheetname = "Clubs";
                    placehold = 13;
                    necessary_group = "Varsity";
                    break;
                case 'S': // Club Sport
                    sheetname = "Clubs";
                    placehold = 13;
                    necessary_group = "Club sport";
                    break;
                case 'p': //POC
                    sheetname = "Mental Health";
                    placehold = 1;
                    necessary_group = "POC";
                    break;
                case 'd': // Depression
                    sheetname = "Mental Health";
                    placehold = 1;
                    necessary_group = "Depression/Anxiety";
                    break;
                case 'a': // academic issue
                    sheetname = "Mental Health";
                    placehold = 1;
                    necessary_group = "Academic Stress";
                    break;
                case 'H': // homesick 
                    sheetname = "Homesickness";
                    placehold = 1;
                    necessary_group = "Homesick";
                    break;
                case 'F': //fitness 
                    sheetname = "FitnessNutrition";
                    placehold = 1;
                    necessary_group = "Fitness";
                    break;
                case 'n': // nutrition
                    sheetname = "FitnessNutrition";
                    placehold = 1;
                    necessary_group = "Nutrition";
                    break;
                case 'C': // CC
                    sheetname = "On-Campus";
                    placehold = 1;
                    necessary_group = "CC";
                    break;
                case 'o': // Foisie
                    sheetname = "On-Campus";
                    placehold = 1;
                    necessary_group = "Foisie";
                    break;
                case 'w': // Workspaces
                    sheetname = "On-Campus";
                    placehold = 1;
                    necessary_group = "Workspaces";
                    break;
                case 'G': // shopping
                    sheetname = "Off-Campus";
                    placehold = 1;
                    necessary_group = "shopping";
                    break;
                case 'O': // Other fun activity
                    sheetname = "Fun Activities";
                    placehold = 1;
                    necessary_group = "Other";
                    break;
                case 'T': // Physical Fun Activity
                    sheetname = "Fun Activities";
                    placehold = 1;
                    necessary_group = "Physical";
                    break;
                case 'k': //Cooking
                    sheetname = "Fun Activities";
                    placehold = 1;
                    necessary_group = "Cooking";
                    break;
                case 'R': // Restaurant
                    sheetname = "Fun Activities";
                    placehold = 1;
                    necessary_group = "Restaurant";
                    break;
                case 'i': // festival
                    sheetname = "Fun Activities";
                    placehold = 1;
                    necessary_group = "Festival";
                    break;
                case 'e': // concerts
                    sheetname = "Fun Activities";
                    placehold = 1;
                    necessary_group = "Converts";
                    break;
                case 'l': // Online Activities
                    sheetname = "Online";
                    placehold = 1;
                    necessary_group = "Activities";
                    break;
                case 'r': // Arts/Crafts
                    sheetname = "Online";
                    placehold = 1;
                    necessary_group = "ArtsandCraft";
                    break;
                case 'M': //games
                    sheetname = "Online";
                    placehold = 1;
                    necessary_group = "games";
                    break;
                case 'L': // Discord Class
                    sheetname = "Discord";
                    placehold = 1;
                    necessary_group = "Class Servers";
                    break;
                case 'j': // Discord Major
                    sheetname = "Discord";
                    placehold = 1;
                    necessary_group = "Majors";
                    break;
                case 'y': // Discord House
                    sheetname = "Discord";
                    placehold = 1;
                    necessary_group = "Housing";
                    break;
                case 'Y': // Discord club sport
                    sheetname = "Discord";
                    placehold = 1;
                    necessary_group = "Clubs";
                    break;
                case 'Z': // Discord Video Game
                    sheetname = "Discord";
                    placehold = 1;
                    necessary_group = "Gaming";
                    break;
                case 'z':
                    sheetname = "Clubs";
                    placehold = 13;
                    necessary_group = "Intramural";
                    break;
            }
            try
            {
                xlapp.Open(excelfile, sheetname);
                using (StreamWriter file = new StreamWriter(@"C:\Users\Zach\Pictures\Hackathon\Survey_Result " + DateTime.Now.ToString("MM-dd-yyyy-hh-mm-ss") + ".txt"))
                { // File.writeline
                    for (int i = 2; i < 200; i++)
                    {
                        if (xlapp.ReadCell01(1, i) == "No Data")
                            break;
                        else if (xlapp.ReadCell01(placehold, i) == necessary_group)
                        {
                            file.WriteLine(xlapp.ReadCell01(1, i));

                            for (int j = 2; j < 50; j++)
                            {
                                if (xlapp.ReadCell01(j, i) == "CHECK")
                                    break;
                                else
                                    file.WriteLine("\t" + xlapp.ReadCell01(j, 1) + ": " + xlapp.ReadCell01(j, i));
                            }
                        }
                    }
                }
                System.Diagnostics.Debug.WriteLine("Done");
            }
            finally
            {
                xlapp.Closed();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            Cursor.Current = Cursors.Default;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label_picture2_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        { 
            Excel xlapp = new Excel();
            string date = monthCalendar1.SelectionRange.Start.ToString();
            if (date.Contains("1/24"))
            {
                sheetname = "Jan24";
            }
            else if (date.Contains("1/25"))
            {
                sheetname = "Jan25";
            }
            else if(date.Contains("1/26"))
            {
                sheetname = "Jan26";
            }
            else if(date.Contains("1/27"))
            {
                sheetname = "Jan27";
            }
            else if(date.Contains("1/28"))
            {
                sheetname = "Jan28";
            }
            else if(date.Contains("1/29"))
            {
                sheetname = "Jan29";
            }
            else
            {
                sheetname = "Jan30";
            }
            

            try
            {
                xlapp.Open(excelfile, sheetname);
                // File.writeline
                this.label1.Text = "Name: " + xlapp.ReadCell01(1, 2) + "\nTime: " + xlapp.ReadCell01(2, 2) + "\nDescription: " + xlapp.ReadCell01(3, 2) + "\nLocation: " + xlapp.ReadCell01(4, 2);
                this.label3.Text = "Name: " + xlapp.ReadCell01(1, 3) + "\nTime: " + xlapp.ReadCell01(2, 3) + "\nDescription: " + xlapp.ReadCell01(3, 3) + "\nLocation: " + xlapp.ReadCell01(4, 3);
                this.label4.Text = "Name: " + xlapp.ReadCell01(1, 4) + "\nTime: " + xlapp.ReadCell01(2, 4) + "\nDescription: " + xlapp.ReadCell01(3, 4) + "\nLocation: " + xlapp.ReadCell01(4, 4);

                System.Diagnostics.Debug.WriteLine("Done");
            }
            finally
            {
                xlapp.Closed();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            if (date.Contains("1/24/2021"))
            {
                this.label1.Visible = true;
                this.label3.Visible = true;
                this.label4.Visible = false;
            }
            else if (date.Contains("1/25/2021"))
            {
                this.label1.Visible = true;
                this.label3.Visible = false;
                this.label4.Visible = false;
            } 
            else if (date.Contains("1/26/2021"))
            {
                this.label1.Visible = true;
                this.label3.Visible = true;
                this.label4.Visible = false;
            }
            else if (date.Contains("1/27/2021"))
            {
                this.label1.Visible = true;
                this.label3.Visible = true;
                this.label4.Visible = false;
            }
            else if (date.Contains("1/28/2021"))
            {
                this.label1.Visible = true;
                this.label3.Visible = false;
                this.label4.Visible = false;
            }
            else if (date.Contains("1/29/2021"))
            {
                this.label1.Visible = true;
                this.label3.Visible = false;
                this.label4.Visible = false;
            }
            
            else if (date.Contains("1/30/2021"))
            {
                this.label1.Visible = false;
                this.label3.Visible = false;
                this.label4.Visible = false;
            }
            
            else
            {
                this.label1.Visible = false;
                this.label3.Visible = false;
                this.label4.Visible = false;
            }
        }

        private void label1_Click_1(object sender, EventArgs e)
        {
             
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void Nav_Events_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void button_delete_Click(object sender, EventArgs e)
        {
            string[] files = Directory.GetFiles(@"C:\Users\Zach\Pictures\Hackathon\");
            foreach (string file in files)
            {
                File.Delete(file);
            }
            linkLabel1.Text = "";
            linkLabel2.Text = "";
            linkLabel3.Text = "";
            linkLabel4.Text = "";

        }

        void load_club_list ()
        {
            sheetname = "Clubs";
            Excel xlapp = new Excel();
            try
            {
                xlapp.Open(excelfile, sheetname);
                using (StreamWriter file = new StreamWriter(@"C:\Users\Zach\Pictures\Hack Support\Clubs" + ".txt"))
                { // File.writeline
                    for (int i = 2; i < 200; i++)
                    {
                        if (xlapp.ReadCell01(1, i) == "No Data")
                            break;
                        else
                            file.WriteLine(xlapp.ReadCell01(1, i));
                        
                    }
                }
                System.Diagnostics.Debug.WriteLine("Done");
            }
            finally
            {
                xlapp.Closed();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            string[] clublist = File.ReadAllLines(@"C:\Users\Zach\Pictures\Hack Support\Clubs" + ".txt");
            for(int i = 0; i < clublist.Length; i++)
            {
                this.dropBox_1.Items.Add(clublist[i]);
            }
        }
        string targetEmail = "";
        private void dropBox_1_TextUpdate(object sender, EventArgs e)
        {

        }

        private void button_email_Click(object sender, EventArgs e)
        {
            Email email = new Email();
            email.SendEmail(targetEmail, dropBox_1.Text);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void dropBox_1_TextChanged(object sender, EventArgs e)
        {
            sheetname = "Clubs";
            necessary_group = dropBox_1.Text;
            Excel xlapp = new Excel();
            try
            {
                xlapp.Open(excelfile, sheetname);
                for (int i = 2; i < 200; i++)
                {
                    if (xlapp.ReadCell01(1, i) == "No Data")
                        break;
                    else if (xlapp.ReadCell01(1, i) == necessary_group)
                        targetEmail = xlapp.ReadCell01(4, i);
                }
            }
            finally
            {
                xlapp.Closed();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            if (targetEmail == "No Data")
                button_email.Enabled = false;
            else
                button_email.Enabled = true;
        }
    }
}
