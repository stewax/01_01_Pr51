using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word_kazakov.Context;

namespace Word_kazakov.Elements
{
    /// <summary>
    /// Логика взаимодействия для Room.xaml
    /// </summary>
    public partial class Room : UserControl
    {
        public Room(int Room)
        {
            InitializeComponent();
            NameRoom.Content = "Квартира номер: " + Room;
            LoadOwner(Room);
        }

        public void LoadOwner(int Room)
        {
            List<OwnerContext> roomOwn = OwnerContext.AllOwners().FindAll(x => x.NumberRoom == Room);
            foreach (OwnerContext room in roomOwn)
            {
                Parent.Children.Add(new Elements.Owner(room));
            }
        }
    }
}
