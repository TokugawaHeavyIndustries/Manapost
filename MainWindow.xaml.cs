using EasyPost;
using EasyPost.Models.API;
using PdfiumViewer;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Net.Http;
using System.Windows;
using System.Windows.Controls;

namespace Manapost
{
    public partial class MainWindow : Window
    {
        private const string ConfigFileName = "config.txt";
        private string apiKey = "";
        private string labelSize = "";
        private string fromName = "";
        private string fromAddress1 = "";
        private string fromAddress2 = "";
        private string fromCity = "";
        private string fromState = "";
        private string fromZip = "";

        bool addressEditor = false;
        bool labelLimitCalculated = false;
        int apiKeyValid = -1;

        int labelCount = 0;


        public MainWindow()
        {
            InitializeComponent();
            LoadConfig();

            WeightLbsEntry.TextChanged += PkgSizeChanged;
            WeightOzEntry.TextChanged += PkgSizeChanged;
            ToTypePicker.SelectionChanged += PkgSizeChanged;
            MachineCheckBox.Checked += PkgSizeChanged;
            MachineCheckBox.Unchecked += PkgSizeChanged;

            ToTypePicker.SelectedIndex = 0;

            if (!string.IsNullOrEmpty(apiKey))
            {
                Task.Run(async () =>
                {
                    await Dispatcher.InvokeAsync(async () =>
                    {
                        await ValidateApiKeyOnStartup();
                    });
                });
            }

            if (PrinterPicker.SelectedItem == null)
            {
                PrinterSelection();
            }
        }

        private void PkgSizeChanged(object sender, EventArgs e)
        {
            QuoteCost();
        }

        private void PkgSizeChanged(object sender, RoutedEventArgs e)
        {
            QuoteCost();
        }

        private void PrinterSelection()
        {
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                PrinterPicker.Items.Add(printer);
            }
        }

        private void LoadConfig()
        {
            var exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string configPath = Path.Combine(exeDirectory, ConfigFileName);
            if (!File.Exists(configPath))
            {
                using (File.Create(configPath)) { }
            }

            var lines = File.ReadAllLines(configPath);
            var config = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line) || !line.Contains('=')) continue;

                var parts = line.Split('=', 2);
                if (parts.Length == 2)
                    config[parts[0].Trim()] = parts[1].Trim();
            }

            config.TryGetValue("PrintDirectly", out var printDirectly);
            config.TryGetValue("Printer", out var printer);
            config.TryGetValue("ApiKey", out apiKey);
            config.TryGetValue("LabelSize", out labelSize);
            config.TryGetValue("FromName", out fromName);
            config.TryGetValue("FromAddress1", out fromAddress1);
            config.TryGetValue("FromAddress2", out fromAddress2);
            config.TryGetValue("FromCity", out fromCity);
            config.TryGetValue("FromState", out fromState);
            config.TryGetValue("FromZip", out fromZip);

            if (config.TryGetValue("AddressType", out var addressType))
            { 

            if (bool.TryParse(addressType, out bool showAddressState))
            {
                    System.Diagnostics.Debug.WriteLine($"AddressType from config: {showAddressState}");
                    AddressToggle.IsChecked = showAddressState;
                AddressToggled(AddressToggle, new RoutedEventArgs());
            }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"AddressType from config: Not Found");
                AddressToggle.IsChecked = true;
                AddressToggled(AddressToggle, new RoutedEventArgs());
            }

                PrinterSelection();
            PrinterPicker.SelectedItem = printer;
            if (string.IsNullOrEmpty(printDirectly))
            {
                printDirectly = "false";
            }
            bool printDirectlyBool = bool.Parse(printDirectly);
            PrintDirectlyCheckBox.IsChecked = printDirectlyBool;
            ApiKeyEntry.Text = apiKey;
            NameEntry.Text = fromName;
            Address1Entry.Text = fromAddress1;
            Address2Entry.Text = fromAddress2;
            CityEntry.Text = fromCity;
            StateEntry.Text = fromState;
            ZipEntry.Text = fromZip;

            if (!string.IsNullOrEmpty(labelSize))
            {
                foreach (ComboBoxItem item in LabelSizePicker.Items)
                {
                    if (item.Content.ToString() == labelSize)
                    {
                        LabelSizePicker.SelectedItem = item;
                        break;
                    }
                }
            }

            GetCurrentBalance();
            LabelsThisMonth();

        }

        private void SavePrinterSettings(object sender, RoutedEventArgs e)
        {
            if (LabelSizePicker.SelectedItem is ComboBoxItem selectedItem)
            {
                UpdateConfigValue("LabelSize", selectedItem.Content.ToString());
            }

            if (PrinterPicker.SelectedItem is string printer)
            {
                UpdateConfigValue("Printer", printer);
            }

            UpdateConfigValue("PrintDirectly", PrintDirectlyCheckBox.IsChecked.ToString());
        }

        private void SaveAPIKey(object sender, RoutedEventArgs e)
        {
            apiKey = ApiKeyEntry.Text?.Trim() ?? "";
            UpdateConfigValue("APIKey", ApiKeyEntry.Text ?? "");
            GetCurrentBalance();
            LabelsThisMonth();
        }

        private void SaveFromAddress(object sender, RoutedEventArgs e)
        {
            UpdateConfigValue("fromName", NameEntry.Text ?? "");
            UpdateConfigValue("fromAddress1", Address1Entry.Text ?? "");
            UpdateConfigValue("fromAddress2", Address2Entry.Text ?? "");
            UpdateConfigValue("fromCity", CityEntry.Text ?? "");
            UpdateConfigValue("fromState", StateEntry.Text ?? "");
            UpdateConfigValue("fromZip", ZipEntry.Text ?? "");
        }

        void UpdateConfigValue(string key, string value)
        {
            var exeDirectory = AppDomain.CurrentDomain.BaseDirectory;

            string configPath = Path.Combine(exeDirectory, ConfigFileName);
            var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();
            bool keyFound = false;

            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                if (string.IsNullOrWhiteSpace(line) || !line.Contains("="))
                    continue;

                var parts = line.Split('=', 2);
                if (parts[0].Trim().Equals(key, StringComparison.OrdinalIgnoreCase))
                {
                    lines[i] = $"{key}={value}";
                    keyFound = true;
                    break;
                }
            }

            if (!keyFound)
            {
                lines.Add($"{key}={value}");
            }

            File.WriteAllLines(configPath, lines);
        }

        private void AddressToggled(object sender, RoutedEventArgs e)
        {
            bool isToType = AddressToggle.IsChecked ?? false;

            ToNameEntry.Visibility = isToType ? Visibility.Visible : Visibility.Collapsed;
            ToAddress1Entry.Visibility = isToType ? Visibility.Visible : Visibility.Collapsed;
            ToAddress2Entry.Visibility = isToType ? Visibility.Visible : Visibility.Collapsed;
            ToCityEntry.Visibility = isToType ? Visibility.Visible : Visibility.Collapsed;
            ToStateEntry.Visibility = isToType ? Visibility.Visible : Visibility.Collapsed;
            ToZipEntry.Visibility = isToType ? Visibility.Visible : Visibility.Collapsed;

            CopyPastableAddress.Visibility = !isToType ? Visibility.Visible : Visibility.Collapsed;

            addressEditor = isToType;

            UpdateConfigValue("AddressType", isToType.ToString().ToLower());
        }

        private async void BuyLabel(object sender, RoutedEventArgs e)
        {
            var client = new EasyPost.Client(new EasyPost.ClientConfiguration(apiKey));

            string toName;
            string toAddress1;
            string toAddress2;
            string toCity;
            string toState;
            string toZip;

            // idk why but addresses with second line crash the program, figure this out 

            if (addressEditor == false)
            {
                string editorText = CopyPastableAddress.Text ?? "";

                string[] lines = editorText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                toName = lines.Length > 0 ? lines[0].Trim() : "";
                toAddress1 = lines.Length > 1 ? lines[1].Trim() : "";
                toAddress2 = "";

                toCity = "";
                toState = "";
                toZip = "";

                if (lines.Length > 2)
                {
                    string cityStateZip = lines[2].Trim();
                    var parts = cityStateZip.Split(',');

                    if (parts.Length == 2)
                    {
                        toCity = parts[0].Trim();

                        var stateZipParts = parts[1].Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);

                        if (stateZipParts.Length >= 2)
                        {
                            toState = stateZipParts[0].Trim();
                            toZip = stateZipParts[1].Trim();
                        }
                    }
                }
            }
            else
            {
                toName = ToNameEntry.Text?.Trim();
                toAddress1 = ToAddress1Entry.Text?.Trim();
                toAddress2 = ToAddress2Entry.Text?.Trim();
                toCity = ToCityEntry.Text?.Trim();
                toState = ToStateEntry.Text?.Trim();
                toZip = ToZipEntry.Text?.Trim();
            }

            bool machinable = !(MachineCheckBox.IsChecked ?? false);
            string labelSize = (LabelSizePicker.SelectedItem as ComboBoxItem)?.Content?.ToString();

            string packageType = (ToTypePicker.SelectedItem as ComboBoxItem)?.Content?.ToString();

            string weightlbs = WeightLbsEntry.Text?.Trim() ?? "0";
            string weightoz = WeightOzEntry.Text?.Trim() ?? "0";

            if (!decimal.TryParse(weightlbs, out decimal lbs))
                lbs = 0;

            if (!decimal.TryParse(weightoz, out decimal oz))
                oz = 0;

            decimal totalOz = (lbs * 16m) + oz;
            string weightTotalOz = totalOz.ToString("0.##");

            var fromAddress = await client.Address.Create(new Dictionary<string, object>
            {
                { "name", fromName },
                { "street1", fromAddress1 },
                { "street2", fromAddress2 },
                { "city", fromCity },
                { "state", fromState },
                { "zip", fromZip }
            });

            var toAddress = await client.Address.Create(new Dictionary<string, object>
            {
                { "name", toName },
                { "street1", toAddress1 },
                { "street2", toAddress2 },
                { "city", toCity },
                { "state", toState },
                { "zip", toZip },
            });

            var parcel = await client.Parcel.Create(new Dictionary<string, object>
            {
                { "predefined_package", packageType },
                { "weight", totalOz } // api expects oz
            });

            var shipment = await client.Shipment.Create(new Dictionary<string, object>
            {
                { "to_address", toAddress },
                { "from_address", fromAddress },
                { "parcel", parcel },
                { "options", new Dictionary<string, object>
                    {
                        { "label_format", "PDF" },     // force pdf for now
                        { "label_size", labelSize },
                        { "machinable", machinable }
                    }
                }
            });

            System.Diagnostics.Debug.WriteLine(weightTotalOz);
            System.Diagnostics.Debug.WriteLine(shipment.Rates);
            var rate = shipment.LowestRate(
                new List<string> { "USPS" },
                new List<string> { "First" }
                );

            var buyParameters = new EasyPost.Parameters.Shipment.Buy(rate);
            var purchasedShipment = await client.Shipment.Buy(shipment.Id, buyParameters);

            System.Diagnostics.Debug.WriteLine($"Purchased Label ID: {purchasedShipment.PostageLabel.Id}");
            System.Diagnostics.Debug.WriteLine($"Purchased Label Format: {purchasedShipment.PostageLabel.LabelFileType}");
            System.Diagnostics.Debug.WriteLine($"Label URL: {purchasedShipment.PostageLabel.LabelUrl}");

            var exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var labelDirectory = Path.Combine(exeDirectory, "Labels");

            // Create directory if it doesn't exist
            if (!Directory.Exists(labelDirectory))
            {
                Directory.CreateDirectory(labelDirectory);
            }

            var postageLabelId = purchasedShipment.PostageLabel.Id;

            string labelUrl = purchasedShipment.PostageLabel.LabelUrl;
            var httpClient = new HttpClient();
            var labelData = await httpClient.GetByteArrayAsync(labelUrl);
            string fileName = $"Label_{postageLabelId}.pdf";
            string filePath = Path.Combine(labelDirectory, fileName);
            await File.WriteAllBytesAsync(filePath, labelData);

            if (PrintDirectlyCheckBox.IsChecked == true)
            {
                PrintDocument(filePath);
            }
            else
            {
                // Open the file with default PDF viewer
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            GetCurrentBalance();
            LabelsThisMonth(true);
        }

        private CancellationTokenSource _debounceCts;

        private async void QuoteCost()
        {
            string weightlbs = WeightLbsEntry.Text?.Trim() ?? "0";
            string weightoz = WeightOzEntry.Text?.Trim() ?? "0";

            if (!decimal.TryParse(weightlbs, out decimal lbs))
                lbs = 0;

            if (!decimal.TryParse(weightoz, out decimal oz))
                oz = 0;

            decimal totalOz = (lbs * 16m) + oz;
            string weightTotalOz = totalOz.ToString("0.##");

            var client = new EasyPost.Client(new EasyPost.ClientConfiguration(apiKey));

            _debounceCts?.Cancel();
            _debounceCts = new CancellationTokenSource();

            string packageType = (ToTypePicker.SelectedItem as ComboBoxItem)?.Content?.ToString();
            bool machinable = !(MachineCheckBox.IsChecked ?? false);

            try
            {
                var fromAddress = await client.Address.Create(new Dictionary<string, object>
                {
                    { "name", "Laura Palmer" },
                    { "street1", "708 Northwestern St" },
                    { "city", "Twin Peaks" },
                    { "state", "WA" },
                    { "zip", "99153" }
                });

                var toAddress = await client.Address.Create(new Dictionary<string, object>
                {
                    { "name", "Special Agent Dale Cooper" },
                    { "street1", "500 Great Northern Highway" },
                    { "street2", "Room 315" },
                    { "city", "Twin Peaks" },
                    { "state", "WA" },
                    { "zip", "99153" }
                });

                var parcel = await client.Parcel.Create(new Dictionary<string, object>
                {
                    { "predefined_package", packageType },
                    { "weight", weightTotalOz } // in ounces
                });

                var shipment = await client.Shipment.Create(new Dictionary<string, object>
                {
                    { "to_address", toAddress },
                    { "from_address", fromAddress },
                    { "parcel", parcel },
                    { "options", new Dictionary<string, object>
                        {
                            { "machinable", machinable }
                        }
                    }
                });

                var rate = shipment.LowestRate(
                    new List<string> { "USPS" },
                    new List<string> { "First" }
                    );

                var rateAmount = rate.Price;

                Cost.Text = rateAmount;
            }
            catch (Exception ex)
            {
                Cost.Text = "Error";
            }
        }

        public async Task PrintDocument(string filePath)
        {

            string selectedprinterName = PrinterPicker.SelectedItem as string;

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");

            try
            {
                await Task.Run(() =>
                {
                    using (var document = PdfDocument.Load(filePath))
                    {
                        using (var printDoc = document.CreatePrintDocument())
                        {
                            printDoc.PrinterSettings = new PrinterSettings
                            {
                                PrinterName = selectedprinterName
                            };

                            printDoc.PrintController = new StandardPrintController(); // Hides the print dialog
                            printDoc.Print();
                        }
                    }
                    ;
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to print PDF: {ex.Message}", ex);
            }
        }

        private void OnPickerSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = (ComboBox)sender;
            int selectedIndex = comboBox.SelectedIndex;
            if (selectedIndex == 1)
            {
                MachineCheckBox.IsEnabled = false;
            }
            else
            {
                MachineCheckBox.IsEnabled = true;
            }
        }

        private async void TestAPIKey(object sender, RoutedEventArgs e)
        {
            apiKey = ApiKeyEntry.Text?.Trim();
            if (LabelSizePicker.SelectedItem is ComboBoxItem selectedItem)
                labelSize = selectedItem.Content.ToString();

            if (string.IsNullOrEmpty(apiKey))
            {
                DynamicLabel.Text = "Please enter an EasyPost API key.";
                return;
            }

            try
            {
                var client = new Client(new ClientConfiguration(apiKey));
                User user = await client.User.RetrieveMe();

                if (user != null && !string.IsNullOrEmpty(user.Id))
                {
                    DynamicLabel.Text = $"Connected to EasyPost as {user.Name}";
                    apiKeyValid = 1;
                    GetCurrentBalance();
                    LabelsThisMonth();
                }
                else
                {
                    DynamicLabel.Text = "API key is invalid or user not found.";
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("production API Key"))
                {
                    DynamicLabel.Text = "Test environment API key entered.";
                    apiKeyValid = 0;
                    LabelsThisMonth();
                    return;
                }
                else
                { 
                    DynamicLabel.Text = $"Failed to connect: {ex.Message}";
                }
            }
        }

        private async Task ValidateApiKeyOnStartup()
        {
            try
            {
                var client = new Client(new ClientConfiguration(apiKey));
                User user = await client.User.RetrieveMe();

                if (user != null && !string.IsNullOrEmpty(user.Id))
                {
                    DynamicLabel.Text = $"Connected to EasyPost as {user.Name}";
                    apiKeyValid = 1;
                    GetCurrentBalance();
                }
                else
                {
                    DynamicLabel.Text = "API key is invalid or user not found.";
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("production API Key"))
                {
                    DynamicLabel.Text = "Test environment API key entered.";
                    apiKeyValid = 0;
                    LabelsThisMonth();
                    return;
                }
                else
                {
                    DynamicLabel.Text = $"Failed to connect: {ex.Message}";
                }
            }
        }

        private async void GetCurrentBalance()
        {
            if (apiKeyValid > 0)
            {
                var client = new Client(new ClientConfiguration(apiKey));
                User user = await client.User.Retrieve();
                string balance = user.Balance;
                balance = balance[..Math.Min(balance.Length, 5)];
                AccountBalance.Text = $"${balance}";
            }
        }

        private void PrinterPicker_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private async void LabelsThisMonth(bool labelPrinted = false)
        {

            string beforeId = null;
            double labelsRemaining;

            if (labelPrinted && labelLimitCalculated)
            {
                labelCount++;
                labelsRemaining = 3000 - labelCount;
                LabelsRemaining.Text = labelsRemaining.ToString();

                Debug.WriteLine($"Total labels this month: {labelCount}");
                Debug.WriteLine($"Labels remaining: {labelsRemaining}");
            }
            if (labelPrinted && !labelLimitCalculated)
            {
               labelCount++;
            }


            if (apiKeyValid >= 0 && labelPrinted == false)
            {
                LabelsRemaining.Text = "Calculating...";
                labelCount = 0;
                labelsRemaining = 0;
                var startOfMonth = new DateTime(DateTime.UtcNow.Year, DateTime.UtcNow.Month, 1);
                var endDate = DateTime.UtcNow;

                var client = new Client(new ClientConfiguration(apiKey));

                while (true)
                {
                    var parameters = new Dictionary<string, object>
                        {
                        { "page_size", 100 },
                        { "start_datetime", startOfMonth.ToString("yyyy-MM-dd'T'00:00:00'Z'") },
                        { "end_datetime", endDate.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'") }
                        };

                    if (!string.IsNullOrEmpty(beforeId))
                        parameters.Add("before_id", beforeId);

                    var shipmentCollection = await client.Shipment.All(parameters);
                    List<Shipment> shipments = shipmentCollection.Shipments;

                    Console.WriteLine($"Fetched {shipments.Count} shipments");

                    if (shipments == null || shipments.Count == 0)
                        break;

                    foreach (var shipment in shipments)
                    {
                        Console.WriteLine($"Shipment ID: {shipment.Id}, CreatedAt: {shipment.CreatedAt}, LabelUrl: {shipment.PostageLabel?.LabelUrl}");

                        if (shipment.PostageLabel != null && !string.IsNullOrWhiteSpace(shipment.PostageLabel.LabelUrl))
                            labelCount++;
                    }

                    beforeId = shipments[^1].Id;
                }

                labelsRemaining = 3000 - labelCount;
                LabelsRemaining.Text = labelsRemaining.ToString();

                Debug.WriteLine($"Total labels this month: {labelCount}");
                Debug.WriteLine($"Labels remaining: {labelsRemaining}");

                labelLimitCalculated = true;
            }
        }

    }
}