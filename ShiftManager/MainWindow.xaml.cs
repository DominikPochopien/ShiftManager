using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace GrafikManager
{
    public class Employee
    {
        public string Name { get; set; } = "";
        public string Department { get; set; } = "";
    }

    public class Role
    {
        public string Name { get; set; } = string.Empty;
        public string ColorHex { get; set; } = string.Empty;
        public string Department { get; set; } = "";
    }

    public class ShiftDef
    {
        public string Hours { get; set; } = string.Empty;
        public string Department { get; set; } = "";
    }

    public class Shift
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string EmployeeName { get; set; } = string.Empty;
        public string RoleName { get; set; } = string.Empty;
        public string RoleColor { get; set; } = string.Empty;
        public string Hours { get; set; } = string.Empty;
        public double HoursCount { get; set; }

        public string TextColor => (RoleColor?.ToUpper() == "#FFFFFF") ? "Black" : "White";
    }

    public class CalendarDayVM
    {
        public DateTime Date { get; set; }
        public string DisplayDate => Date == DateTime.MinValue ? "" : Date.ToString("dd.MM");
        public string BackgroundColor { get; set; } = "#1E1E1E";
        public string DateColor { get; set; } = "#555";
        public ObservableCollection<Shift> Shifts { get; set; } = new ObservableCollection<Shift>();
    }

    public class AppData
    {
        public List<Employee> Employees { get; set; } = new List<Employee>();
        public List<Role> Roles { get; set; } = new List<Role>();
        public List<ShiftDef> ShiftDefs { get; set; } = new List<ShiftDef>();
        public Dictionary<string, List<Shift>> AllShifts { get; set; } = new Dictionary<string, List<Shift>>();
        public List<string> Calendars { get; set; } = new List<string> { "Sala", "Kuchnia", "FoodTruck" };
        public List<string> Departments { get; set; } = new List<string>();
    }

    public partial class MainWindow : Window
    {
        private AppData _data = new AppData();
        private DateTime _currentMonth;
        private string? _currentCalendar = null;
        private string _savePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "GrafikManager");

        private readonly string[] PaletteColors = {
            "#D32F2F", "#C62828", "#B71C1C", "#FF1744", "#AD1457", "#880E4F", "#C2185B", "#E91E63",
            "#7B1FA2", "#6A1B9A", "#4A148C", "#AA00FF", "#6200EA", "#304FFE", "#1A237E", "#283593",
            "#1976D2", "#1565C0", "#0D47A1", "#0277BD", "#01579B", "#00838F", "#006064", "#0097A7",
            "#00796B", "#004D40", "#2E7D32", "#1B5E20", "#33691E", "#558B2F", "#827717", "#9E9D24",
            "#F57F17", "#FF6F00", "#E65100", "#BF360C", "#D84315", "#4E342E", "#3E2723", "#5D4037",
            "#424242", "#616161", "#455A64", "#37474F", "#263238", "#212121", "#757575", "#546E7A",
            "#DD2C00", "#FF3D00", "#2962FF", "#00BFA5", "#00C853", "#64DD17", "#AEEA00", "#FFD600",
            "#311B92", "#1A237E", "#004D40", "#1B5E20", "#BF360C", "#3E2723", "#263238", "#FFFFFF"
        };

        private string _activeEmployee = "";
        private Role? _activeRole = null;
        private string _activeHours = "";
        private string _tempRoleColor = "#FFFFFF";

        // Inicjalizuje ustawienia początkowe aplikacji, foldery i interfejs.
        public MainWindow()
        {
            InitializeComponent();
            InitializeApp();
        }

        // Przygotowuje dane startowe, datę oraz odświeża główne widoki.
        private void InitializeApp()
        {
            Directory.CreateDirectory(_savePath);
            _currentMonth = DateTime.Now;
            RightPanelContainer.Visibility = Visibility.Hidden;
            UpdateMonthDisplay();
            RefreshToolPanels();
        }

        // Cofa kalendarz o jeden miesiąc i odświeża siatkę.
        private void BtnPrevMonth_Click(object sender, RoutedEventArgs e)
        {
            _currentMonth = _currentMonth.AddMonths(-1);
            UpdateMonthDisplay();
        }

        // Przesuwa kalendarz o jeden miesiąc do przodu i odświeża siatkę.
        private void BtnNextMonth_Click(object sender, RoutedEventArgs e)
        {
            _currentMonth = _currentMonth.AddMonths(1);
            UpdateMonthDisplay();
        }

        // Aktualizuje etykietę tekstową bieżącego miesiąca na ekranie.
        private void UpdateMonthDisplay()
        {
            TxtCurrentMonth.Text = _currentMonth.ToString("MMMM yyyy").ToUpper();
            if (_currentCalendar != null) GenerateCalendarGrid();
        }

        // Oblicza dni miesiąca, puste komórki i generuje strukturę siatki kalendarza.
        private void GenerateCalendarGrid()
        {
            if (_currentCalendar == null) return;

            var daysList = new List<CalendarDayVM>();
            int daysInMonth = DateTime.DaysInMonth(_currentMonth.Year, _currentMonth.Month);
            DateTime firstDay = new DateTime(_currentMonth.Year, _currentMonth.Month, 1);

            int startOffset = (int)firstDay.DayOfWeek - 1;
            if (startOffset < 0) startOffset = 6;

            for (int i = 0; i < startOffset; i++)
            {
                daysList.Add(new CalendarDayVM { Date = DateTime.MinValue, BackgroundColor = "#141414", DateColor = "Transparent" });
            }

            for (int i = 1; i <= daysInMonth; i++)
            {
                DateTime date = new DateTime(_currentMonth.Year, _currentMonth.Month, i);
                var dayVM = new CalendarDayVM
                {
                    Date = date,
                    BackgroundColor = date == DateTime.Today ? "#2D2D30" : "#1E1E1E",
                    DateColor = date.DayOfWeek == DayOfWeek.Sunday ? "#FF5252" : "#888"
                };

                string key = GetKey(date, _currentCalendar);
                if (_data.AllShifts.ContainsKey(key))
                {
                    foreach (var s in _data.AllShifts[key]) dayVM.Shifts.Add(s);
                }
                daysList.Add(dayVM);
            }

            int totalDaysSoFar = startOffset + daysInMonth;
            int remainingCells = 7 - (totalDaysSoFar % 7);

            if (remainingCells < 7)
            {
                for (int i = 0; i < remainingCells; i++)
                {
                    daysList.Add(new CalendarDayVM { Date = DateTime.MinValue, BackgroundColor = "#141414", DateColor = "Transparent" });
                }
            }

            CalendarItems.ItemsSource = daysList;
            CalculateStats();
        }

        // Tworzy unikalny klucz słownika dla konkretnego dnia i wybranego kalendarza.
        private string GetKey(DateTime date, string calendar) => $"{date:yyyy-MM-dd}|{calendar}";

        // Czyści i ponownie ładuje przyciski dla działów, pracowników, ról i godzin.
        private void RefreshToolPanels()
        {
            PanelCalendarsList.Children.Clear();
            foreach (var cal in _data.Calendars)
            {
                var btn = CreateToolButton(cal, cal == _currentCalendar, "#673AB7");
                btn.Click += (s, e) => {
                    _currentCalendar = cal;
                    _activeEmployee = "";
                    _activeRole = null;
                    _activeHours = "";
                    RightPanelContainer.Visibility = Visibility.Visible;
                    UpdateActiveStatus();
                    GenerateCalendarGrid();
                    RefreshToolPanels();
                };
                btn.ContextMenu = CreateDeleteMenu(cal, (s, e_menu) => {
                    if (_data.Calendars.Count <= 1) return;
                    if (MessageBox.Show($"Usunąć dział '{cal}'? Usunie to również wszystkich przypisanych do niego pracowników, role i godziny.", "Potwierdź", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        _data.Calendars.Remove(cal);
                        _data.Employees.RemoveAll(x => x.Department == cal);
                        _data.Roles.RemoveAll(x => x.Department == cal);
                        _data.ShiftDefs.RemoveAll(x => x.Department == cal);

                        if (_currentCalendar == cal)
                        {
                            _currentCalendar = null; _activeEmployee = ""; _activeRole = null; _activeHours = "";
                            RightPanelContainer.Visibility = Visibility.Hidden; UpdateActiveStatus();
                        }
                        GenerateCalendarGrid(); RefreshToolPanels();
                    }
                });
                PanelCalendarsList.Children.Add(btn);
            }

            PanelEmployeesList.Children.Clear();
            if (_currentCalendar != null)
            {
                foreach (var emp in _data.Employees.Where(e => e.Department == _currentCalendar))
                {
                    var btn = CreateToolButton(emp.Name, emp.Name == _activeEmployee, "#2D2D30");
                    btn.Click += (s, e) => { _activeEmployee = emp.Name; UpdateActiveStatus(); RefreshToolPanels(); };
                    btn.ContextMenu = CreateDeleteMenu(emp, (s, e_menu) => {
                        if (MessageBox.Show($"Usunąć: {emp.Name}?", "Potwierdź", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            _data.Employees.Remove(emp);
                            if (_activeEmployee == emp.Name) _activeEmployee = "";
                            RefreshToolPanels();
                        }
                    });
                    PanelEmployeesList.Children.Add(btn);
                }
            }

            PanelRolesList.Children.Clear();
            if (_currentCalendar != null)
            {
                foreach (var role in _data.Roles.Where(r => r.Department == _currentCalendar))
                {
                    var btn = CreateToolButton(role.Name, _activeRole != null && role.Name == _activeRole.Name, role.ColorHex);
                    btn.Click += (s, e) => { _activeRole = role; UpdateActiveStatus(); RefreshToolPanels(); };
                    btn.ContextMenu = CreateDeleteMenu(role, (s, e_menu) => {
                        if (MessageBox.Show($"Usunąć rolę: {role.Name}?", "Potwierdź", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            _data.Roles.Remove(role);
                            if (_activeRole == role) _activeRole = null;
                            RefreshToolPanels();
                        }
                    });
                    PanelRolesList.Children.Add(btn);
                }
            }

            PanelHoursList.Children.Clear();
            if (_currentCalendar != null)
            {
                foreach (var shiftDef in _data.ShiftDefs.Where(s => s.Department == _currentCalendar))
                {
                    var btn = CreateToolButton(shiftDef.Hours, shiftDef.Hours == _activeHours, "#2D2D30");
                    btn.Click += (s, e) => { _activeHours = shiftDef.Hours; UpdateActiveStatus(); RefreshToolPanels(); };
                    btn.ContextMenu = CreateDeleteMenu(shiftDef, (s, e_menu) => {
                        if (MessageBox.Show($"Usunąć: {shiftDef.Hours}?", "Potwierdź", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            _data.ShiftDefs.Remove(shiftDef);
                            if (_activeHours == shiftDef.Hours) _activeHours = "";
                            RefreshToolPanels();
                        }
                    });
                    PanelHoursList.Children.Add(btn);
                }
            }
        }

        // Generuje dynamiczny przycisk interfejsu z przypisanymi właściwościami wizualnymi.
        private Button CreateToolButton(string content, bool isActive, string colorHex)
        {
            bool isWhiteRole = colorHex?.ToUpper() == "#FFFFFF";

            var btn = new Button
            {
                Content = content,
                Margin = new Thickness(2),
                Padding = new Thickness(5, 2, 5, 2),
                Background = IsHexColor(colorHex) ? (SolidColorBrush)new BrushConverter().ConvertFrom(colorHex)! : Brushes.Gray,
                Foreground = isWhiteRole ? Brushes.Black : Brushes.White
            };

            if (isActive)
            {
                btn.BorderThickness = new Thickness(2);
                btn.BorderBrush = isWhiteRole ? Brushes.Black : Brushes.White;
                btn.FontWeight = FontWeights.Bold;
            }
            else
            {
                btn.BorderThickness = new Thickness(1);
                btn.BorderBrush = Brushes.Transparent;
                if (colorHex != "#2D2D30" && !isWhiteRole) btn.Opacity = 0.7;
            }
            return btn;
        }

        // Weryfikuje, czy przekazany ciąg znaków jest prawidłowym kodem koloru HEX.
        private bool IsHexColor(string hex) => !string.IsNullOrEmpty(hex) && hex.StartsWith("#");

        // Odświeża etykiety informacyjne pokazujące aktualnie zaznaczone elementy.
        private void UpdateActiveStatus()
        {
            TxtActiveCalendar.Text = $"Dział: {(_currentCalendar ?? "-")}";
            TxtActiveCalendar.Foreground = _currentCalendar != null ? Brushes.MediumPurple : Brushes.Gray;

            TxtActiveEmp.Text = $"Pracownik: {(string.IsNullOrEmpty(_activeEmployee) ? "-" : _activeEmployee)}";
            TxtActiveEmp.Foreground = !string.IsNullOrEmpty(_activeEmployee) ? Brushes.LightGreen : Brushes.White;

            TxtActiveRole.Text = $"Rola: {(_activeRole?.Name ?? "-")}";
            TxtActiveRole.Foreground = _activeRole != null ? (SolidColorBrush)new BrushConverter().ConvertFrom(_activeRole.ColorHex)! : Brushes.White;

            TxtActiveHours.Text = $"Godziny: {(_activeHours ?? "-")}";
            TxtActiveHours.Foreground = _activeHours != null ? Brushes.LightBlue : Brushes.White;
        }

        // Dodaje wpisanego pracownika do bieżącego działu po naciśnięciu Enter.
        private void InputEmployee_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && !string.IsNullOrWhiteSpace(InputEmployee.Text))
            {
                if (_currentCalendar == null)
                {
                    MessageBox.Show("Aby dodać pracownika, musisz najpierw wybrać i kliknąć w DZIAŁ (np. Kuchnia, Sala)!", "Brak wybranego działu", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string newName = InputEmployee.Text.Trim();
                if (_data.Employees.Any(x => x.Name.Equals(newName, StringComparison.OrdinalIgnoreCase) && x.Department == _currentCalendar))
                {
                    MessageBox.Show($"Pracownik '{newName}' jest już dodany do działu '{_currentCalendar}'!", "Duplikat");
                    return;
                }

                _data.Employees.Add(new Employee { Name = newName, Department = _currentCalendar });
                InputEmployee.Text = "";
                RefreshToolPanels();
            }
        }

        // Wyświetla paletę kolorów pozwalającą na wybranie barwy dla nowej roli.
        private void BtnRoleColor_Click(object sender, RoutedEventArgs e)
        {
            Button? btn = sender as Button;
            ContextMenu colorPalette = new ContextMenu();
            UniformGrid grid = new UniformGrid { Columns = 8, Width = 280 };

            foreach (var hex in PaletteColors)
            {
                Border colorSquare = new Border
                {
                    Background = (SolidColorBrush)new BrushConverter().ConvertFrom(hex)!,
                    Width = 30,
                    Height = 30,
                    Margin = new Thickness(1),
                    CornerRadius = new CornerRadius(3),
                    Cursor = Cursors.Hand,
                    BorderBrush = hex.ToUpper() == "#FFFFFF" ? Brushes.Gray : Brushes.Transparent,
                    BorderThickness = new Thickness(1)
                };

                colorSquare.MouseLeftButtonDown += (s, ev) =>
                {
                    _tempRoleColor = hex;
                    if (btn != null) btn.Background = colorSquare.Background;
                    colorPalette.IsOpen = false;
                };
                grid.Children.Add(colorSquare);
            }

            MenuItem wrapperItem = new MenuItem { Header = grid };
            ControlTemplate rawTemplate = new ControlTemplate(typeof(MenuItem));
            FrameworkElementFactory cp = new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.ContentProperty, new TemplateBindingExtension(MenuItem.HeaderProperty));
            rawTemplate.VisualTree = cp;
            wrapperItem.Template = rawTemplate;

            colorPalette.Items.Add(wrapperItem);
            colorPalette.PlacementTarget = btn;
            colorPalette.IsOpen = true;
        }

        // Zapisuje nową rolę przypisaną do aktywnego działu po wciśnięciu Enter.
        private void InputRole_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && !string.IsNullOrWhiteSpace(InputRole.Text))
            {
                if (_currentCalendar == null)
                {
                    MessageBox.Show("Aby dodać rolę, musisz najpierw wybrać DZIAŁ!", "Brak wybranego działu", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string newRole = InputRole.Text.Trim();
                if (_data.Roles.Any(x => x.Name.Equals(newRole, StringComparison.OrdinalIgnoreCase) && x.Department == _currentCalendar)) return;

                _data.Roles.Add(new Role { Name = newRole, ColorHex = _tempRoleColor, Department = _currentCalendar });
                InputRole.Text = "";
                RefreshToolPanels();
            }
        }

        // Rejestruje nowy przedział godzin dla aktywnego kalendarza po zatwierdzeniu Enterem.
        private void InputHours_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && !string.IsNullOrWhiteSpace(InputHours.Text))
            {
                if (_currentCalendar == null)
                {
                    MessageBox.Show("Aby dodać godziny, musisz najpierw wybrać DZIAŁ!", "Brak wybranego działu", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string newHours = InputHours.Text.Trim();
                if (_data.ShiftDefs.Any(x => x.Hours == newHours && x.Department == _currentCalendar)) return;

                _data.ShiftDefs.Add(new ShiftDef { Hours = newHours, Department = _currentCalendar });
                InputHours.Text = "";
                RefreshToolPanels();
            }
        }

        // Dodaje nowy dział operacyjny do listy kalendarzy po naciśnięciu klawisza Enter.
        private void InputCalendarName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && !string.IsNullOrWhiteSpace(InputCalendarName.Text))
            {
                string newName = InputCalendarName.Text.Trim();
                if (!_data.Calendars.Contains(newName))
                {
                    _data.Calendars.Add(newName);
                    _currentCalendar = newName;
                    _activeEmployee = ""; _activeRole = null; _activeHours = "";
                    RightPanelContainer.Visibility = Visibility.Visible;
                    InputCalendarName.Text = "";
                    UpdateActiveStatus(); GenerateCalendarGrid(); RefreshToolPanels();
                }
            }
        }

        // Rozpoznaje kliknięcie w konkretny dzień na siatce i inicjuje przypisanie zmiany.
        private void Day_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Tag is DateTime date && date != DateTime.MinValue)
            {
                if (!string.IsNullOrEmpty(_activeEmployee) && _activeRole != null && !string.IsNullOrEmpty(_activeHours))
                    AddShiftToDate(date, _activeEmployee, _activeRole, _activeHours);
                else
                    MessageBox.Show("Zaznacz wszystkie 3 narzędzia (Pracownik, Rola, Godziny) z lewego panelu!", "Brak");
            }
        }

        // Oblicza długość dyżuru i dodaje wpis o zmianie do słownika dla wskazanego dnia.
        private void AddShiftToDate(DateTime date, string empName, Role role, string hours)
        {
            if (_currentCalendar == null) return;
            string key = GetKey(date, _currentCalendar);
            if (!_data.AllShifts.ContainsKey(key)) _data.AllShifts[key] = new List<Shift>();
            if (_data.AllShifts[key].Any(x => x.EmployeeName == empName)) return;

            double hCount = 0;
            try
            {
                var parts = hours.Split('-');
                if (parts.Length == 2 && double.TryParse(parts[0], out double s) && double.TryParse(parts[1], out double end))
                {
                    hCount = end - s;
                    if (hCount < 0) hCount += 24;
                }
            }
            catch { }

            _data.AllShifts[key].Add(new Shift
            {
                EmployeeName = empName,
                RoleName = role.Name,
                RoleColor = role.ColorHex,
                Hours = hours,
                HoursCount = hCount
            });
            GenerateCalendarGrid();
        }

        // Usuwa wskazaną zmianę (dyżur) pracownika wybraną z menu kontekstowego.
        private void MenuRemoveShift_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem item && item.Tag is Shift shiftToRemove)
            {
                foreach (var key in _data.AllShifts.Keys)
                {
                    var found = _data.AllShifts[key].FirstOrDefault(x => x.Id == shiftToRemove.Id);
                    if (found != null)
                    {
                        _data.AllShifts[key].Remove(found);
                        break;
                    }
                }
                GenerateCalendarGrid();
            }
        }

        // Podlicza i wyświetla łączną liczbę godzin przepracowanych w bieżącym miesiącu przez personel.
        private void CalculateStats()
        {
            PanelStats.Children.Clear();
            if (_currentCalendar == null) return;

            var stats = new Dictionary<string, double>();
            int daysInMonth = DateTime.DaysInMonth(_currentMonth.Year, _currentMonth.Month);

            for (int i = 1; i <= daysInMonth; i++)
            {
                string key = GetKey(new DateTime(_currentMonth.Year, _currentMonth.Month, i), _currentCalendar);
                if (_data.AllShifts.ContainsKey(key))
                {
                    foreach (var s in _data.AllShifts[key])
                    {
                        if (!stats.ContainsKey(s.EmployeeName)) stats[s.EmployeeName] = 0;
                        stats[s.EmployeeName] += s.HoursCount;
                    }
                }
            }

            foreach (var kvp in stats.OrderByDescending(x => x.Value))
            {
                var grid = new Grid { Margin = new Thickness(0, 2, 0, 2) };
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                grid.Children.Add(new TextBlock { Text = kvp.Key, Foreground = Brushes.LightGray });
                var txtHours = new TextBlock { Text = $"{kvp.Value}h", FontWeight = FontWeights.Bold, Foreground = Brushes.White };
                Grid.SetColumn(txtHours, 1);
                grid.Children.Add(txtHours);
                PanelStats.Children.Add(grid);
            }
        }

        // Bezpowrotnie usuwa wszystkie wpisane dotychczas zmiany po zatwierdzeniu przez użytkownika.
        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Wyczyścić CAŁY GRAFIK?", "Potwierdź", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                _data.AllShifts.Clear();
                GenerateCalendarGrid();
            }
        }

        // Formatuje i eksportuje podstawowe struktury bazy (działy, role, pracownicy) do pliku tekstowego.
        private void BtnSaveEmployees_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var lines = new List<string> { "---EMPLOYEES---" };
                lines.AddRange(_data.Employees.Select(x => $"{x.Name}|{x.Department}"));
                lines.Add("---ROLES---");
                lines.AddRange(_data.Roles.Select(x => $"{x.Name}|{x.ColorHex}|{x.Department}"));
                lines.Add("---SHIFTS---");
                lines.AddRange(_data.ShiftDefs.Select(x => $"{x.Hours}|{x.Department}"));

                File.WriteAllLines(Path.Combine(_savePath, "baza_danych.txt"), lines);
                MessageBox.Show("Zapisano!");
            }
            catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
        }

        // Analizuje zawartość pliku bazowego i ładuje konfigurację działów do programu.
        private void BtnLoadEmployees_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string path = Path.Combine(_savePath, "baza_danych.txt");
                if (File.Exists(path))
                {
                    var lines = File.ReadAllLines(path);
                    string mode = "";
                    _data.Employees.Clear(); _data.Roles.Clear(); _data.ShiftDefs.Clear();

                    foreach (var line in lines)
                    {
                        if (line.StartsWith("---")) { mode = line; continue; }

                        var p = line.Split('|');
                        if (mode == "---EMPLOYEES---")
                        {
                            if (p.Length == 2) _data.Employees.Add(new Employee { Name = p[0], Department = p[1] });
                            else _data.Employees.Add(new Employee { Name = line, Department = _data.Calendars.FirstOrDefault() ?? "" });
                        }
                        else if (mode == "---ROLES---")
                        {
                            if (p.Length == 3) _data.Roles.Add(new Role { Name = p[0], ColorHex = p[1], Department = p[2] });
                            else if (p.Length == 2) _data.Roles.Add(new Role { Name = p[0], ColorHex = p[1], Department = _data.Calendars.FirstOrDefault() ?? "" });
                        }
                        else if (mode == "---SHIFTS---")
                        {
                            if (p.Length == 2) _data.ShiftDefs.Add(new ShiftDef { Hours = p[0], Department = p[1] });
                            else _data.ShiftDefs.Add(new ShiftDef { Hours = line, Department = _data.Calendars.FirstOrDefault() ?? "" });
                        }
                    }
                    RefreshToolPanels();
                    MessageBox.Show("Wczytano!");
                }
            }
            catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
        }

        // Serializuje i zapisuje układ grafiku ze wszystkimi zmianami do lokalnego pliku JSON.
        private void BtnSaveSchedule_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                File.WriteAllText(Path.Combine(_savePath, "grafik.json"), JsonSerializer.Serialize(_data.AllShifts));
                MessageBox.Show("Zapisano!");
            }
            catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
        }

        // Pobiera zawartość pliku JSON z grafikiem i nadpisuje aktualnie wyświetlany widok.
        private void BtnLoadSchedule_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string path = Path.Combine(_savePath, "grafik.json");
                if (File.Exists(path))
                {
                    _data.AllShifts = JsonSerializer.Deserialize<Dictionary<string, List<Shift>>>(File.ReadAllText(path))!;
                    if (_currentCalendar != null) GenerateCalendarGrid();
                    RefreshToolPanels();
                    MessageBox.Show("Wczytano!");
                }
            }
            catch (Exception ex) { MessageBox.Show("Błąd: " + ex.Message); }
        }

        // Generuje standardowe menu kontekstowe (Prawy Przycisk Myszy) obsługujące operację usuwania.
        private ContextMenu CreateDeleteMenu(object target, RoutedEventHandler deleteHandler)
        {
            ContextMenu menu = new ContextMenu();
            MenuItem deleteItem = new MenuItem { Header = "Usuń", Tag = target };
            deleteItem.Click += deleteHandler;
            menu.Items.Add(deleteItem);
            return menu;
        }

        // Eksportuje dane dotyczące wybranego działu do bogato formatowanego arkusza kalkulacyjnego.
        private void BtnExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_currentCalendar == null)
            {
                MessageBox.Show("Najpierw wybierz dział (kalendarz), który chcesz wyeksportować!", "Błąd", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var employeesToExport = _data.Employees.Where(emp => emp.Department == _currentCalendar).Select(emp => emp.Name).ToList();

            foreach (var kvp in _data.AllShifts)
            {
                if (kvp.Key.EndsWith("|" + _currentCalendar))
                {
                    foreach (var shift in kvp.Value)
                    {
                        if (!employeesToExport.Contains(shift.EmployeeName))
                            employeesToExport.Add(shift.EmployeeName);
                    }
                }
            }

            employeesToExport.Sort();

            if (employeesToExport.Count == 0)
            {
                MessageBox.Show("Brak pracowników i zmian w tym dziale. Nie ma czego eksportować!", "Informacja", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"Grafik_{_currentCalendar}_{_currentMonth:MM_yyyy}.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new ClosedXML.Excel.XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(_currentCalendar);
                        int daysInMonth = DateTime.DaysInMonth(_currentMonth.Year, _currentMonth.Month);
                        int colIndex = 2;

                        Dictionary<int, int> dayColumns = new Dictionary<int, int>();
                        List<int> weeklySumColumns = new List<int>();

                        for (int i = 1; i <= daysInMonth; i++)
                        {
                            DateTime date = new DateTime(_currentMonth.Year, _currentMonth.Month, i);
                            dayColumns[i] = colIndex;
                            colIndex++;

                            if (date.DayOfWeek == DayOfWeek.Sunday || i == daysInMonth)
                            {
                                weeklySumColumns.Add(colIndex);
                                colIndex++;
                            }
                        }
                        int totalSumColIndex = colIndex;

                        worksheet.Cell(1, 1).Value = $"GRAFIK: {_currentCalendar.ToUpper()} - {_currentMonth:MMMM yyyy}".ToUpper();
                        worksheet.Range(1, 1, 1, totalSumColIndex).Merge();
                        worksheet.Cell(1, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(1, 1).Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                        worksheet.Cell(1, 1).Style.Font.Bold = true;
                        worksheet.Cell(1, 1).Style.Font.FontSize = 16;
                        worksheet.Row(1).Height = 25;

                        worksheet.Cell(2, 1).Value = "PRACOWNIK";
                        worksheet.Range(2, 1, 3, 1).Merge();
                        worksheet.Cell(2, 1).Style.Font.Bold = true;
                        worksheet.Cell(2, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.White;
                        worksheet.Cell(2, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.Black;
                        worksheet.Cell(2, 1).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                        worksheet.Cell(2, 1).Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;
                        worksheet.Cell(2, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(2, 1).Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;

                        worksheet.Row(3).Height = 85;

                        for (int i = 1; i <= daysInMonth; i++)
                        {
                            int c = dayColumns[i];
                            DateTime date = new DateTime(_currentMonth.Year, _currentMonth.Month, i);

                            worksheet.Cell(2, c).Value = i;
                            worksheet.Cell(2, c).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(2, c).Style.Font.Bold = true;
                            worksheet.Cell(2, c).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                            worksheet.Cell(2, c).Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;

                            worksheet.Cell(3, c).Value = date.ToString("dddd", new System.Globalization.CultureInfo("pl-PL")).ToUpper();
                            worksheet.Cell(3, c).Style.Alignment.TextRotation = 90;
                            worksheet.Cell(3, c).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(3, c).Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                            worksheet.Cell(3, c).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                            worksheet.Cell(3, c).Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;
                            worksheet.Cell(3, c).Style.Font.FontSize = 9;

                            if (date.DayOfWeek == DayOfWeek.Sunday)
                            {
                                worksheet.Cell(2, c).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#A0A0A0");
                                worksheet.Cell(3, c).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#A0A0A0");
                            }
                            else if (date.DayOfWeek == DayOfWeek.Saturday)
                            {
                                worksheet.Cell(2, c).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#E0E0E0");
                                worksheet.Cell(3, c).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#E0E0E0");
                            }
                            else if (i % 2 == 0)
                            {
                                worksheet.Cell(2, c).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#FFF9C4");
                                worksheet.Cell(3, c).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#FFF9C4");
                            }
                            else
                            {
                                worksheet.Cell(2, c).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.White;
                                worksheet.Cell(3, c).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.White;
                            }
                        }

                        foreach (int c in weeklySumColumns)
                        {
                            worksheet.Cell(2, c).Value = "SUMA TYG.";
                            worksheet.Range(2, c, 3, c).Merge();
                            worksheet.Cell(2, c).Style.Alignment.TextRotation = 90;
                            worksheet.Cell(2, c).Style.Font.Bold = true;
                            worksheet.Cell(2, c).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#FF3333");
                            worksheet.Cell(2, c).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(2, c).Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                            worksheet.Cell(2, c).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                            worksheet.Cell(2, c).Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;
                            worksheet.Cell(2, c).Style.Font.FontSize = 9;
                        }

                        worksheet.Cell(2, totalSumColIndex).Value = "SUMA MIES.";
                        worksheet.Range(2, totalSumColIndex, 3, totalSumColIndex).Merge();
                        worksheet.Cell(2, totalSumColIndex).Style.Alignment.TextRotation = 90;
                        worksheet.Cell(2, totalSumColIndex).Style.Font.Bold = true;
                        worksheet.Cell(2, totalSumColIndex).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#FF3333");
                        worksheet.Cell(2, totalSumColIndex).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(2, totalSumColIndex).Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                        worksheet.Cell(2, totalSumColIndex).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                        worksheet.Cell(2, totalSumColIndex).Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;

                        int currentRow = 4;

                        foreach (var empName in employeesToExport)
                        {
                            double monthlyTotal = 0;
                            double weeklyTotal = 0;

                            worksheet.Row(currentRow).Height = 14.5;
                            worksheet.Row(currentRow + 1).Height = 14.5;
                            worksheet.Row(currentRow + 2).Height = 14.5;

                            worksheet.Range(currentRow, 1, currentRow + 1, 1).Merge();
                            worksheet.Cell(currentRow, 1).Value = empName;
                            worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                            worksheet.Cell(currentRow, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.White;
                            worksheet.Cell(currentRow, 1).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                            worksheet.Cell(currentRow, 1).Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;
                            worksheet.Cell(currentRow, 1).Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;

                            var spacerEmp = worksheet.Cell(currentRow + 2, 1);
                            spacerEmp.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.White;
                            spacerEmp.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                            spacerEmp.Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;

                            for (int i = 1; i <= daysInMonth; i++)
                            {
                                DateTime date = new DateTime(_currentMonth.Year, _currentMonth.Month, i);
                                string key = GetKey(date, _currentCalendar);
                                int c = dayColumns[i];

                                var topCell = worksheet.Cell(currentRow, c);
                                var bottomCell = worksheet.Cell(currentRow + 1, c);
                                var spacerCell = worksheet.Cell(currentRow + 2, c);

                                ClosedXML.Excel.XLColor columnBaseColor;
                                if (date.DayOfWeek == DayOfWeek.Sunday) columnBaseColor = ClosedXML.Excel.XLColor.FromHtml("#A0A0A0");
                                else if (date.DayOfWeek == DayOfWeek.Saturday) columnBaseColor = ClosedXML.Excel.XLColor.FromHtml("#E0E0E0");
                                else if (i % 2 == 0) columnBaseColor = ClosedXML.Excel.XLColor.FromHtml("#FFF9C4");
                                else columnBaseColor = ClosedXML.Excel.XLColor.White;

                                topCell.Style.Fill.BackgroundColor = columnBaseColor;
                                bottomCell.Style.Fill.BackgroundColor = columnBaseColor;
                                spacerCell.Style.Fill.BackgroundColor = columnBaseColor;

                                topCell.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                topCell.Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;
                                bottomCell.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                bottomCell.Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;
                                spacerCell.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                spacerCell.Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;

                                topCell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                                topCell.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                                bottomCell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                                bottomCell.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;

                                if (_data.AllShifts.ContainsKey(key))
                                {
                                    var shift = _data.AllShifts[key].FirstOrDefault(s => s.EmployeeName == empName);
                                    if (shift != null)
                                    {
                                        topCell.Value = shift.Hours.Replace("-", ";");
                                        topCell.Style.Font.FontSize = 10;
                                        topCell.Style.Font.Bold = true;

                                        bottomCell.Value = shift.HoursCount;
                                        bottomCell.Style.Font.FontSize = 10;
                                        bottomCell.Style.Font.Bold = true;

                                        weeklyTotal += shift.HoursCount;
                                        monthlyTotal += shift.HoursCount;

                                        if (!string.IsNullOrEmpty(shift.RoleColor) && shift.RoleColor.ToUpper() != "#FFFFFF")
                                        {
                                            try
                                            {
                                                var roleColor = ClosedXML.Excel.XLColor.FromHtml(shift.RoleColor);
                                                topCell.Style.Fill.BackgroundColor = roleColor;
                                                bottomCell.Style.Fill.BackgroundColor = roleColor;
                                            }
                                            catch { }
                                        }

                                        topCell.Style.Font.FontColor = ClosedXML.Excel.XLColor.Black;
                                        bottomCell.Style.Font.FontColor = ClosedXML.Excel.XLColor.Black;
                                    }
                                }

                                if (date.DayOfWeek == DayOfWeek.Sunday || i == daysInMonth)
                                {
                                    int sumCol = c + 1;
                                    var sumRange = worksheet.Range(currentRow, sumCol, currentRow + 1, sumCol);
                                    sumRange.Merge();
                                    sumRange.Value = weeklyTotal > 0 ? weeklyTotal.ToString() : "";
                                    sumRange.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                                    sumRange.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                                    sumRange.Style.Font.Bold = true;
                                    sumRange.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                    sumRange.Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;
                                    sumRange.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#FF3333");

                                    var spacerSum = worksheet.Cell(currentRow + 2, sumCol);
                                    spacerSum.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#FF3333");
                                    spacerSum.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                                    spacerSum.Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;

                                    weeklyTotal = 0;
                                }
                            }

                            var monthSumRange = worksheet.Range(currentRow, totalSumColIndex, currentRow + 1, totalSumColIndex);
                            monthSumRange.Merge();
                            monthSumRange.Value = monthlyTotal > 0 ? monthlyTotal.ToString() : "";
                            monthSumRange.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                            monthSumRange.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                            monthSumRange.Style.Font.Bold = true;
                            monthSumRange.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                            monthSumRange.Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;
                            monthSumRange.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#FF3333");

                            var spacerMonthSum = worksheet.Cell(currentRow + 2, totalSumColIndex);
                            spacerMonthSum.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#FF3333");
                            spacerMonthSum.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                            spacerMonthSum.Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.Black;

                            currentRow += 3;
                        }

                        worksheet.Column(1).Width = 18;

                        for (int i = 1; i <= daysInMonth; i++)
                        {
                            worksheet.Column(dayColumns[i]).Width = 4.3;
                        }
                        foreach (int c in weeklySumColumns)
                        {
                            worksheet.Column(c).Width = 4.3;
                        }
                        worksheet.Column(totalSumColIndex).Width = 4.3;

                        worksheet.SheetView.FreezeRows(3);
                        worksheet.SheetView.FreezeColumns(1);

                        workbook.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Plik Excel został wygenerowany pomyślnie!", "Sukces", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Błąd: {ex.Message}", "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}