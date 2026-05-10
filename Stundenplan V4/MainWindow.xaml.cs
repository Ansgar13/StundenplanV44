using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Stundenplan_V2
{
    public partial class MainWindow : Window
    {
        private string excelPfad = "";

        private StundenplanInput input;

        private readonly StundenplanService service =
            new StundenplanService(new OrToolsSolver());

        // label = "OhneTausch_1", "OhneTausch_2", "Tausch_5+7_1" usw.
        // blocks = die für diese Lösung gültigen Blöcke (ggf. mit getauschten Lehrern)
        private List<(int quality, int badUnits, int[,] belegung, string label, List<UnterrichtsBlock> blocks)> letzteSolutions = new();

        public MainWindow()
        {
            InitializeComponent();
        }

        // =====================================================
        // LOG
        // =====================================================
        public void Log(string text)
        {
            TxtLog.AppendText(text + Environment.NewLine);
            TxtLog.ScrollToEnd();
        }

        // =====================================================
        // BUTTON 1 – ZEITWÜNSCHE EINLESEN
        // =====================================================
        private void BtnZeitwuensche_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(excelPfad))
            {
                MessageBox.Show("Bitte zuerst Excel-Datei laden (Button 2).");
                return;
            }

            var dlgTxt = new OpenFileDialog();
            dlgTxt.Filter = "Textdateien (*.txt)|*.txt";
            dlgTxt.InitialDirectory = System.IO.Path.GetDirectoryName(excelPfad);

            if (dlgTxt.ShowDialog() != true)
                return;

            try
            {
                ZeitwunschExporter.ErzeugeZeitWL(excelPfad, dlgTxt.FileName);
                TxtStatus.Text = "ZeitWL und ZeitWK erzeugt.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler:\n" + ex.Message);
            }
        }

        // =====================================================
        // BUTTON 2 – EXCEL EINLESEN
        // =====================================================
        private void BtnPfad_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Filter = "Excel Dateien (*.xlsx)|*.xlsx";

            if (dlg.ShowDialog() == true)
            {
                excelPfad = dlg.FileName;
                input = ExcelLoader.Lade(excelPfad);
                TxtStatus.Text = "Excel erfolgreich eingelesen.";
            }
        }

        // =====================================================
        // BUTTON 3 – STUNDENPLANERSTELLUNG
        // =====================================================
        private void BtnSchritt2_Click(object sender, RoutedEventArgs e)
        {
            if (input == null)
            {
                MessageBox.Show("Bitte zuerst Excel-Datei laden.");
                return;
            }

            var statusFenster = new Window
            {
                Title = "Bitte warten",
                Width = 300,
                Height = 120,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.ToolWindow,
                Topmost = true
            };

            var txt = new System.Windows.Controls.TextBlock
            {
                Text = "Engine sucht ...",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            statusFenster.Content = txt;
            statusFenster.Show();

            // UI sofort rendern, bevor der Solver den Thread blockiert
            System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(
                System.Windows.Threading.DispatcherPriority.Render,
                new Action(() => { }));

            Log("Starte Solver...");

            var solutions = service.Generate(input, Log, out string debug);

            statusFenster.Close();


            if (solutions.Count == 0)
            {
                MessageBox.Show("Keine Lösung gefunden – weder mit noch ohne Tausch.\n\n" + debug);
                TxtStatus.Text = "Planung fehlgeschlagen.";
                return;
            }

            letzteSolutions = solutions.ToList();

            Log($"Lösungen gefunden: {letzteSolutions.Count}");
            foreach (var l in letzteSolutions)
                Log($"  [{l.label}] Qualität: {l.quality}, BadUnits: {l.badUnits}");

            // In Excel schreiben
            SchreibeInExcel(solutions);
            SchreibeRanking(solutions);

            // Diagnose-Tabelle für alle Lösungen
            try
            {
                var diagnoseDaten = letzteSolutions
                    .Select(sol => (
                        sol.label,
                        LehrerDiagnose.Berechne(
                            sol.belegung,
                            sol.blocks,
                            input.Slots,
                            input.LehrerStammdaten,
                            input.StrafeHohlstunde,
                            input.StrafeDoppelHohlstunde,
                            input.StrafeDreifachHohlstunde,
                            input.StrafeStdFolge)))
                    .ToList();

                LehrerDiagnose.Exportiere(excelPfad, diagnoseDaten);
                Log("Diagnose-Tabelle erstellt.");
            }
            catch (Exception ex)
            {
                Log($"Diagnose-Fehler: {ex.Message}");
            }

            TxtStatus.Text = "Stundenverteilung abgeschlossen.";
        }

        // =====================================================
        // BUTTON 4 – LEHRERPLÄNE
        // =====================================================
        private void BtnLehrerplaene_Click(object sender, RoutedEventArgs e)
        {
            if (input == null)
            {
                MessageBox.Show("Bitte zuerst Excel-Datei laden (Button 2).");
                return;
            }

            // UNrPlan in letzteSolutions laden falls noch nicht vorhanden
            if (!letzteSolutions.Any(s => s.label == "UNrPlan"))
                BtnUnrPlan_Click(null, null);

            var verfügbareLösungen = letzteSolutions.Count > 0
                ? letzteSolutions
                : LadeLösungenAusExcel();

            if (verfügbareLösungen.Count == 0)
            {
                MessageBox.Show("Keine Lösungen verfügbar – bitte zuerst Stundenplan erstellen (Button 3) " +
                                "oder Lösungen in der Excel-Tabelle vorhanden.");
                return;
            }

            LöscheAlteSheets(excelPfad, "Lehrerpläne_");

            foreach (var sol in verfügbareLösungen)
            {
                SetzeLoesungInSlots(sol.belegung);
                LehrerplanGenerator.Erzeuge(excelPfad, sol.blocks, input.Slots, sol.label);
            }

            TxtStatus.Text = "Lehrerpläne für alle Lösungen erzeugt.";
        }

        // =====================================================
        // BUTTON 5 – KLASSENPLÄNE
        // =====================================================
        private void BtnKlassenplaene_Click(object sender, RoutedEventArgs e)
        {
            if (input == null)
            {
                MessageBox.Show("Bitte zuerst Excel-Datei laden (Button 2).");
                return;
            }

            // UNrPlan in letzteSolutions laden falls noch nicht vorhanden
            if (!letzteSolutions.Any(s => s.label == "UNrPlan"))
                BtnUnrPlan_Click(null, null);

            var verfügbareLösungen = letzteSolutions.Count > 0
                ? letzteSolutions
                : LadeLösungenAusExcel();

            if (verfügbareLösungen.Count == 0)
            {
                MessageBox.Show("Keine Lösungen verfügbar – bitte zuerst Stundenplan erstellen (Button 3) " +
                                "oder Lösungen in der Excel-Tabelle vorhanden.");
                return;
            }

            LöscheAlteSheets(excelPfad, "Klassenpläne_");

            foreach (var sol in verfügbareLösungen)
            {
                SetzeLoesungInSlots(sol.belegung);
                KlassenplanGenerator.Erzeuge(excelPfad, sol.blocks, input.Slots, sol.label);
            }

            TxtStatus.Text = "Klassenpläne für alle Lösungen erzeugt.";
        }

        // =====================================================
        // BUTTON 5b – KLASSENPLÄNE NUR EF / Q1 / Q2
        // =====================================================
        private void BtnKlassenplaeneOberstufe_Click(object sender, RoutedEventArgs e)
        {
            if (input == null)
            {
                MessageBox.Show("Bitte zuerst Excel-Datei laden (Button 2).");
                return;
            }

            // UNrPlan in letzteSolutions laden falls noch nicht vorhanden
            if (!letzteSolutions.Any(s => s.label == "UNrPlan"))
                BtnUnrPlan_Click(null, null);

            var verfügbareLösungen = letzteSolutions.Count > 0
                ? letzteSolutions
                : LadeLösungenAusExcel();

            if (verfügbareLösungen.Count == 0)
            {
                MessageBox.Show("Keine Lösungen verfügbar – bitte zuerst Stundenplan erstellen (Button 3) " +
                                "oder Lösungen in der Excel-Tabelle vorhanden.");
                return;
            }

            var oberstufenFilter = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase) { "EF", "Q1", "Q2" };

            LöscheAlteSheets(excelPfad, "Klassenpläne_");

            foreach (var sol in verfügbareLösungen)
            {
                SetzeLoesungInSlots(sol.belegung);
                KlassenplanGenerator.Erzeuge(
                    excelPfad, sol.blocks, input.Slots, sol.label,
                    oberstufenFilter);
            }

            TxtStatus.Text = "Klassenpläne EF/Q1/Q2 erzeugt.";
        }
        private void BtnUnrPlan_Click(object sender, RoutedEventArgs e)
        {
            bool automatisch = sender == null;

            if (input == null)
            {
                if (!automatisch)
                    MessageBox.Show("Bitte zuerst Excel-Datei laden (Button 2).");
                return;
            }

            try
            {
                // Bestehenden UNr-Plan aus Excel lesen
                int[,] belegung = LadeUnrPlanAusExcel();

                if (belegung == null)
                {
                    if (!automatisch)
                        MessageBox.Show("Kein UNr-Plan gefunden. Bitte zuerst die Tabelle 'Unr-Plan' befüllen.");
                    return;
                }

                // Bewerten
                var bewertung = PlanBewertung.Berechne(
                    belegung,
                    input.Blocks,
                    input.Slots,
                    input.GewichtFrüheDoppel,
                    input.GewichtSpäteDoppel,
                    input.GewichtSpätePädEinheiten,
                    input.StrafeHohlstunde,
                    input.StrafeDoppelHohlstunde,
                    input.StrafeDreifachHohlstunde,
                    input.StrafeEinzelstunde,
                    input.StrafeSpäteLkStunden,
                    input.StrafeHauptfachSpät,
                    input.HauptfachSpätAnteilProzent);

                var unrPlan = (bewertung.Quality, bewertung.BadUnits, belegung, "UNrPlan", input.Blocks);

                // In letzteSolutions eintragen (alte UNrPlan-Einträge ersetzen)
                letzteSolutions.RemoveAll(s => s.label == "UNrPlan");
                letzteSolutions.Add(unrPlan);

                // In Lösungen-Tabelle eintragen
                SchreibeInExcel(letzteSolutions);
                SchreibeRanking(letzteSolutions);

                // Diagnose-Tabelle aktualisieren inkl. UNrPlan
                try
                {
                    var diagnoseDaten = letzteSolutions
                        .Select(sol => (
                            sol.label,
                            LehrerDiagnose.Berechne(
                                sol.belegung,
                                sol.blocks,
                                input.Slots,
                                input.LehrerStammdaten,
                                input.StrafeHohlstunde,
                                input.StrafeDoppelHohlstunde,
                                input.StrafeDreifachHohlstunde,
                                input.StrafeStdFolge)))
                        .ToList();
                    LehrerDiagnose.Exportiere(excelPfad, diagnoseDaten);
                }
                catch { /* Diagnose-Fehler ignorieren */ }

                Log($"UNr-Plan bewertet: Qualität={bewertung.Quality}, " +
                    $"FrüheDoppel={bewertung.Early}, SpäteDoppel={bewertung.Late}, " +
                    $"BadUnits={bewertung.BadUnits}");

                if (!automatisch)
                    TxtStatus.Text = "UNr-Plan in Lösungen und SolverRanking eingetragen.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler:\n" + ex.Message);
            }
        }

        // =====================================================
        // BUTTON 8 – PLAN PRÜFEN
        // Prüft den UNrPlan auf Constraint-Verletzungen
        // Kann auch ohne vorherigen Solver-Lauf ausgeführt werden
        // =====================================================
        private void BtnPlanPrüfen_Click(object sender, RoutedEventArgs e)
        {
            if (input == null)
            {
                MessageBox.Show("Bitte zuerst Excel-Datei laden (Button 2).");
                return;
            }

            try
            {
                // Belegung immer aus UNrPlan lesen – erst aus letzteSolutions, dann aus Excel
                int[,] belegung = null;

                // UNrPlan aus letzteSolutions suchen
                var unrPlanSol = letzteSolutions.FirstOrDefault(s => s.label == "UNrPlan");
                if (unrPlanSol.belegung != null)
                {
                    belegung = unrPlanSol.belegung;
                }
                else
                {
                    // Aus Excel lesen
                    belegung = LadeUnrPlanAusLösungsTabelle();
                    if (belegung == null)
                        belegung = LadeUnrPlanAusExcel();
                    if (belegung == null)
                    {
                        MessageBox.Show("Kein UNr-Plan gefunden. Bitte zuerst UNr-Plan erzeugen (Button 6) " +
                                        "oder Stundenplan erstellen (Button 3).");
                        return;
                    }
                }

                var verletzungen = PlanValidator.Prüfe(
                    belegung,
                    input.Blocks,
                    input.Slots,
                    input.GrossePausen);

                PlanValidator.SchreibeTabelle(excelPfad, verletzungen);

                if (verletzungen.Count == 0)
                    Log("✓ Keine Constraint-Verletzungen gefunden.");
                else
                    Log($"⚠️ {verletzungen.Count} Verletzungen gefunden – siehe Tabelle 'Verletzungen'.");

                TxtStatus.Text = "Plan-Prüfung abgeschlossen.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // =====================================================
        // BUTTON 9 – PLAN VERBESSERN
        // =====================================================
        private void BtnPlanVerbessern_Click(object sender, RoutedEventArgs e)
        {
            if (input == null)
            {
                MessageBox.Show("Bitte zuerst Excel-Datei laden (Button 2).");
                return;
            }

            var verfügbareLösungen = letzteSolutions.Count > 0
                ? letzteSolutions
                : LadeLösungenAusExcel();

            if (verfügbareLösungen.Count == 0)
            {
                MessageBox.Show("Keine Lösungen verfügbar – bitte zuerst Stundenplan erstellen (Button 3).");
                return;
            }

            var labels = verfügbareLösungen.Select(s => s.label).ToList();
            var dialog = new VerbesserungsDialog(labels) { Owner = this };

            if (dialog.ShowDialog() != true)
                return;

            var optionen = dialog.Optionen;
            int lösungsIdx = dialog.GewählteLösungsIndex;
            bool alsNeu = dialog.AlsNeueLösung;

            var gewählteLösung = verfügbareLösungen[lösungsIdx];

            Log($"Starte Verbesserung von '{gewählteLösung.label}'...");
            TxtStatus.Text = "Verbesserung läuft...";

            try
            {
                var ergebnis = PlanVerbesserung.Verbessere(
                    gewählteLösung.belegung,
                    gewählteLösung.blocks,
                    input.Slots,
                    input,
                    optionen,
                    Log);

                if (ergebnis.Verbesserung <= 0)
                {
                    Log($"Keine Verbesserung gefunden (Qualität bleibt {ergebnis.AusgangsQualität}).");
                    TxtStatus.Text = "Keine Verbesserung gefunden.";
                    return;
                }

                string neuesLabel = alsNeu
                    ? gewählteLösung.label + "_verbessert"
                    : gewählteLösung.label;

                var verbesserteLösung = (
                    ergebnis.EndQualität,
                    gewählteLösung.badUnits,
                    ergebnis.BesteBelegung,
                    neuesLabel,
                    gewählteLösung.blocks);

                if (alsNeu)
                {
                    letzteSolutions.Add(verbesserteLösung);
                }
                else
                {
                    int idx = letzteSolutions.FindIndex(s => s.label == gewählteLösung.label);
                    if (idx >= 0)
                        letzteSolutions[idx] = verbesserteLösung;
                    else
                        letzteSolutions.Add(verbesserteLösung);
                }

                SchreibeInExcel(letzteSolutions);
                SchreibeRanking(letzteSolutions);

                Log($"✓ Verbesserung abgeschlossen: {ergebnis.AusgangsQualität} → {ergebnis.EndQualität} " +
                    $"(+{ergebnis.Verbesserung})");
                TxtStatus.Text = $"Plan verbessert: Qualität {ergebnis.AusgangsQualität} → {ergebnis.EndQualität}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler bei der Verbesserung:\n" + ex.Message);
                Log($"Fehler: {ex.Message}");
            }
        }

        private int[,] LadeUnrPlanAusLösungsTabelle()
        {
            using var wb = new XLWorkbook(excelPfad);
            if (!wb.Worksheets.Any(ws => ws.Name == "Lösungen"))
                return null;

            var sheet = wb.Worksheet("Lösungen");
            var headerRow = sheet.Row(1);

            // UNrPlan-Spalte suchen
            int maxCol = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 2;
            int unrPlanCol = -1;
            for (int col = 3; col <= maxCol; col++)
            {
                if (headerRow.Cell(col).GetString().Trim() == "UNrPlan")
                {
                    unrPlanCol = col;
                    break;
                }
            }

            if (unrPlanCol == -1) return null;

            int S = input.Slots.Count;
            int B = input.Blocks.Count;

            var unrZuIdx = new Dictionary<int, int>();
            for (int b = 0; b < B; b++)
                unrZuIdx[input.Blocks[b].UNr] = b;

            var slotLookup = new Dictionary<string, int>();
            for (int s = 0; s < S; s++)
                slotLookup[$"{input.Slots[s].WTag}_{input.Slots[s].Stunde}"] = s;

            var belegung = new int[B, S];
            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;

            for (int row = 2; row <= lastRow; row++)
            {
                string wtag = sheet.Cell(row, 1).GetString().Trim();
                if (!int.TryParse(sheet.Cell(row, 2).GetString(), out int stunde))
                    continue;

                string slotKey = $"{wtag}_{stunde}";
                if (!slotLookup.TryGetValue(slotKey, out int sIdx)) continue;

                string zelle = sheet.Cell(row, unrPlanCol).GetString().Trim();
                if (string.IsNullOrEmpty(zelle)) continue;

                foreach (var part in zelle.Split(','))
                {
                    if (int.TryParse(part.Trim(), out int unr) &&
                        unrZuIdx.TryGetValue(unr, out int bIdx))
                        belegung[bIdx, sIdx] = 1;
                }
            }

            return belegung;
        }

        // =====================================================
        // BUTTON 7 – KLASSE FIXIEREN
        // Überträgt die Slots einer Klasse aus einer Lösung
        // in die Tabelle "Fix UNrn" (ohne bestehende zu löschen)
        // =====================================================
        private void BtnKlasseFixieren_Click(object sender, RoutedEventArgs e)
        {
            if (input == null)
            {
                MessageBox.Show("Bitte zuerst Excel-Datei laden (Button 2).");
                return;
            }

            // Lösungen aus Excel lesen falls kein Solver-Lauf vorhanden
            var verfügbareLösungen = letzteSolutions.Count > 0
                ? letzteSolutions
                : LadeLösungenAusExcel();

            if (verfügbareLösungen.Count == 0)
            {
                MessageBox.Show("Keine Lösungen verfügbar – bitte zuerst Stundenplan erstellen (Button 3) " +
                                "oder Lösungen in der Excel-Tabelle vorhanden.");
                return;
            }

            // Alle verfügbaren Klassen und Fächer aus den Blöcken extrahieren
            var alleKlassen = input.Blocks
                .SelectMany(b => b.Teile.SelectMany(t => t.Klassen))
                .Distinct().OrderBy(k => k).ToList();

            var alleFächer = input.Blocks
                .SelectMany(b => b.Teile.Select(t => t.Fach.Trim()))
                .Where(f => !string.IsNullOrEmpty(f))
                .Distinct().OrderBy(f => f).ToList();

            var labels = verfügbareLösungen.Select(s => s.label).ToList();
            var dialog = new KlasseFixierenDialog(labels, alleKlassen, alleFächer)
                { Owner = this };

            if (dialog.ShowDialog() != true)
                return;

            var klassen   = dialog.GewählteKlassen;
            var fächer    = dialog.GewählteFächer;
            int lösungsNr = dialog.GewählteLösungsIndex;

            var sol      = verfügbareLösungen[lösungsNr];
            var belegung = sol.belegung;
            var blocks   = sol.blocks;

            // Blöcke suchen die eine der Klassen ODER eines der Fächer enthalten
            var trefferBlöcke = new List<int>();
            for (int b = 0; b < blocks.Count; b++)
            {
                bool klassenTreffer = klassen.Count > 0 &&
                    blocks[b].Teile.Any(t =>
                        t.Klassen.Any(k => klassen.Contains(k)));

                bool fachTreffer = fächer.Count > 0 &&
                    blocks[b].Teile.Any(t =>
                        fächer.Any(f => f.Equals(t.Fach.Trim(),
                            StringComparison.OrdinalIgnoreCase)));

                if (klassenTreffer || fachTreffer)
                    trefferBlöcke.Add(b);
            }

            if (trefferBlöcke.Count == 0)
            {
                MessageBox.Show("Keine passenden Blöcke gefunden.");
                return;
            }

            // Belegte Slots → UNrn sammeln
            var fixEinträge = new Dictionary<int, List<int>>();
            for (int s = 0; s < input.Slots.Count; s++)
            {
                foreach (int b in trefferBlöcke)
                {
                    if (belegung[b, s] == 1)
                    {
                        if (!fixEinträge.ContainsKey(s))
                            fixEinträge[s] = new List<int>();
                        fixEinträge[s].Add(blocks[b].UNr);
                    }
                }
            }

            if (fixEinträge.Count == 0)
            {
                MessageBox.Show($"Keine belegten Slots in '{sol.label}'.");
                return;
            }

            SchreibeFixUNrn(fixEinträge);

            var teile = new List<string>();
            if (klassen.Count > 0) teile.Add($"Klassen: {string.Join(", ", klassen)}");
            if (fächer.Count > 0)  teile.Add($"Fächer: {string.Join(", ", fächer)}");
            string beschreibung = string.Join(" | ", teile);

            TxtStatus.Text = $"{beschreibung} aus '{sol.label}' in Fix UNrn eingetragen.";
            Log($"Fixiert ({beschreibung}): {fixEinträge.Count} Slots, " +
                $"{fixEinträge.Values.Sum(v => v.Count)} Einträge.");
        }

        // =====================================================
        // LÖSUNGEN AUS EXCEL-TABELLE LESEN
        // Liest alle Lösungs-Spalten aus der "Lösungen"-Tabelle
        // =====================================================
        private List<(int quality, int badUnits, int[,] belegung, string label, List<UnterrichtsBlock> blocks)>
            LadeLösungenAusExcel()
        {
            var result = new List<(int, int, int[,], string, List<UnterrichtsBlock>)>();

            using var wb = new XLWorkbook(excelPfad);
            if (!wb.Worksheets.Any(ws => ws.Name == "Lösungen"))
                return result;

            var sheet = wb.Worksheet("Lösungen");
            var headerRow = sheet.Row(1);

            // Spaltennamen lesen (ab Spalte 3)
            int maxCol = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 2;
            var spaltenLabels = new Dictionary<int, string>();
            for (int col = 3; col <= maxCol; col++)
            {
                string label = headerRow.Cell(col).GetString().Trim();
                if (!string.IsNullOrEmpty(label))
                    spaltenLabels[col] = label;
            }

            if (spaltenLabels.Count == 0)
                return result;

            int S = input.Slots.Count;
            int B = input.Blocks.Count;

            // UNr → Block-Index Lookup
            var unrZuIdx = new Dictionary<int, int>();
            for (int b = 0; b < input.Blocks.Count; b++)
                unrZuIdx[input.Blocks[b].UNr] = b;

            // Slot-Lookup: WTag+Stunde → Slot-Index
            var slotLookup = new Dictionary<string, int>();
            for (int s = 0; s < input.Slots.Count; s++)
                slotLookup[$"{input.Slots[s].WTag}_{input.Slots[s].Stunde}"] = s;

            foreach (var kv in spaltenLabels)
            {
                int col = kv.Key;
                string label = kv.Value;

                // Leere Labels überspringen
                if (string.IsNullOrEmpty(label)) continue;

                var belegung = new int[B, S];

                // Zeilen durchgehen (ab Zeile 2)
                int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
                for (int row = 2; row <= lastRow; row++)
                {
                    string wtag = sheet.Cell(row, 1).GetString().Trim();
                    if (!int.TryParse(sheet.Cell(row, 2).GetString(), out int stunde))
                        continue;

                    string slotKey = $"{wtag}_{stunde}";
                    if (!slotLookup.TryGetValue(slotKey, out int s))
                        continue;

                    string zellWert = sheet.Cell(row, col).GetString().Trim();
                    if (string.IsNullOrEmpty(zellWert)) continue;

                    foreach (var teil in zellWert.Split(','))
                    {
                        if (int.TryParse(teil.Trim(), out int unr) &&
                            unrZuIdx.TryGetValue(unr, out int b))
                            belegung[b, s] = 1;
                    }
                }

                result.Add((0, 0, belegung, label, input.Blocks));
            }

            return result;
        }

        // =====================================================
        // FIX UNRN SCHREIBEN
        // Trägt neue UNrn in "Fix UNrn" ein ohne bestehende
        // Einträge zu löschen oder zu überschreiben.
        // =====================================================
        private void SchreibeFixUNrn(Dictionary<int, List<int>> neueEinträge)
        {
            using var wb = new XLWorkbook(excelPfad);

            IXLWorksheet sheet;
            if (wb.Worksheets.Any(ws => ws.Name == "Fix UNrn"))
                sheet = wb.Worksheet("Fix UNrn");
            else
            {
                sheet = wb.Worksheets.Add("Fix UNrn");
                sheet.Cell(1, 1).Value = "WTag";
                sheet.Cell(1, 2).Value = "Stunde";
            }

            foreach (var kv in neueEinträge)
            {
                int slotIdx = kv.Key;
                var neueUnrn = kv.Value;

                string wtag   = input.Slots[slotIdx].WTag;
                int    stunde = input.Slots[slotIdx].Stunde;

                // Bestehende Zeile für diesen Slot suchen
                IXLRow zielZeile = null;
                foreach (var row in sheet.RowsUsed().Skip(1))
                {
                    if (row.Cell(1).GetString().Trim() == wtag &&
                        row.Cell(2).GetString().Trim() == stunde.ToString())
                    {
                        zielZeile = row;
                        break;
                    }
                }

                if (zielZeile == null)
                {
                    // Neue Zeile am Ende anfügen
                    int neueZeile = sheet.LastRowUsed()?.RowNumber() + 1 ?? 2;
                    sheet.Cell(neueZeile, 1).Value = wtag;
                    sheet.Cell(neueZeile, 2).Value = stunde;
                    zielZeile = sheet.Row(neueZeile);
                }

                // Bestehende UNrn in dieser Zeile sammeln
                var vorhandeneUnrn = new HashSet<int>();
                int letzteCol = zielZeile.LastCellUsed()?.Address.ColumnNumber ?? 2;
                for (int col = 3; col <= letzteCol; col++)
                {
                    if (int.TryParse(zielZeile.Cell(col).GetString(), out int vorh))
                        vorhandeneUnrn.Add(vorh);
                }

                // Nur neue UNrn hinzufügen die noch nicht vorhanden sind
                int nächsteCol = letzteCol + 1;

                foreach (int unr in neueUnrn)
                {
                    if (!vorhandeneUnrn.Contains(unr))
                    {
                        zielZeile.Cell(nächsteCol).Value = unr;
                        vorhandeneUnrn.Add(unr);
                        nächsteCol++;
                    }
                }
            }

            wb.Save();
        }

        // =====================================================
        // TOP-PLÄNE IN TABELLE 2 SCHREIBEN
        // =====================================================
        private void SchreibeInExcel(
            List<(int quality, int badUnits, int[,] belegung, string label, List<UnterrichtsBlock> blocks)> solutions)
        {
            using var workbook = new XLWorkbook(excelPfad);
            var sheet = workbook.Worksheet("Lösungen");

            sheet.Cell(1, 1).Value = "WTag";
            sheet.Cell(1, 2).Value = "Stunde";

            for (int p = 0; p < Math.Min(10, solutions.Count); p++)
                sheet.Cell(1, 3 + p).Value = solutions[p].label;

            for (int s = 0; s < input.Slots.Count; s++)
            {
                sheet.Cell(s + 2, 1).Value = input.Slots[s].WTag;
                sheet.Cell(s + 2, 2).Value = input.Slots[s].Stunde;

                for (int p = 0; p < Math.Min(10, solutions.Count); p++)
                {
                    var belegung = solutions[p].belegung;
                    var unrList = new List<int>();

                    for (int b = 0; b < input.Blocks.Count; b++)
                        if (belegung[b, s] == 1)
                            unrList.Add(input.Blocks[b].UNr);

                    sheet.Cell(s + 2, 3 + p).Value = string.Join(", ", unrList);
                }
            }

            int qualRow = input.Slots.Count + 3;
            sheet.Cell(qualRow, 1).Value = "Qualität";

            for (int p = 0; p < solutions.Count; p++)
                sheet.Cell(qualRow, 3 + p).Value = solutions[p].quality;

            workbook.Save();
        }

        private void SchreibeRanking(
            List<(int quality, int badUnits, int[,] belegung, string label, List<UnterrichtsBlock> blocks)> solutions)
        {
            using var workbook = new XLWorkbook(excelPfad);

            if (workbook.Worksheets.Any(ws => ws.Name == "SolverRanking"))
                workbook.Worksheet("SolverRanking").Delete();

            var sheet = workbook.Worksheets.Add("SolverRanking");

            sheet.Cell(1, 1).Value  = "Plan";
            sheet.Cell(1, 2).Value  = "Label";
            sheet.Cell(1, 3).Value  = "Qualität";
            sheet.Cell(1, 4).Value  = "frühe Doppel";
            sheet.Cell(1, 5).Value  = "späte Doppel";
            sheet.Cell(1, 6).Value  = "päd. Einheiten spät";
            sheet.Cell(1, 7).Value  = "Hohlstunden";
            sheet.Cell(1, 8).Value  = "Doppelhohlstunden";
            sheet.Cell(1, 9).Value  = "Dreifachhohlstunden";
            sheet.Cell(1, 10).Value = "Einzelstunden";
            sheet.Cell(1, 11).Value = "späte LK-Stunden";
            sheet.Cell(1, 12).Value = "Hauptfach zu spät";
            sheet.Cell(1, 13).Value = "Details späte päd. Einheiten";
            sheet.Row(1).Style.Font.Bold = true;

            for (int p = 0; p < solutions.Count; p++)
            {
                var bewertung = PlanBewertung.Berechne(
                    solutions[p].belegung,
                    solutions[p].blocks,
                    input.Slots,
                    input.GewichtFrüheDoppel,
                    input.GewichtSpäteDoppel,
                    input.GewichtSpätePädEinheiten,
                    input.StrafeHohlstunde,
                    input.StrafeDoppelHohlstunde,
                    input.StrafeDreifachHohlstunde,
                    input.StrafeEinzelstunde,
                    input.StrafeSpäteLkStunden,
                    input.StrafeHauptfachSpät,
                    input.HauptfachSpätAnteilProzent);

                sheet.Cell(p + 2, 1).Value  = p + 1;
                sheet.Cell(p + 2, 2).Value  = solutions[p].label;
                sheet.Cell(p + 2, 3).Value  = bewertung.Quality;
                sheet.Cell(p + 2, 4).Value  = bewertung.Early;
                sheet.Cell(p + 2, 5).Value  = bewertung.Late;
                sheet.Cell(p + 2, 6).Value  = bewertung.BadUnits;
                sheet.Cell(p + 2, 7).Value  = bewertung.Hohlstunden;
                sheet.Cell(p + 2, 8).Value  = bewertung.DoppelHohlstunden;
                sheet.Cell(p + 2, 9).Value  = bewertung.DreifachHohlstunden;
                sheet.Cell(p + 2, 10).Value = bewertung.Einzelstunden;
                sheet.Cell(p + 2, 11).Value = bewertung.SpäteLkStunden;
                sheet.Cell(p + 2, 12).Value = bewertung.HauptfachSpätÜberschuss;
                sheet.Cell(p + 2, 13).Value = string.Join("\n", bewertung.Details);
                sheet.Cell(p + 2, 13).Style.Alignment.WrapText = true;
            }

            sheet.Columns().AdjustToContents();
            workbook.Save();
        }

        // =====================================================
        // UNR-PLAN AUS EXCEL LADEN
        // =====================================================
        private int[,] LadeUnrPlanAusExcel()
        {
            int B = input.Blocks.Count;
            int S = input.Slots.Count;
            int[,] belegung = new int[B, S];

            using var wb = new XLWorkbook(excelPfad);

            if (!wb.Worksheets.Any(ws => ws.Name == "Unr-Plan"))
                return null;

            var sheet = wb.Worksheet("Unr-Plan");

            for (int s = 0; s < S; s++)
            {
                int col = 3;
                while (true)
                {
                    var cell = sheet.Cell(s + 2, col);
                    if (cell.IsEmpty()) break;

                    int unr = cell.GetValue<int>();
                    for (int b = 0; b < input.Blocks.Count; b++)
                        if (input.Blocks[b].UNr == unr)
                            belegung[b, s] = 1;

                    col++;
                }
            }

            return belegung;
        }

        private (int quality, int badUnits, int[,] belegung, string label, List<UnterrichtsBlock> blocks) BewerteUnrPlan()
        {
            int[,] belegung = LadeUnrPlanAusExcel();
            var b = PlanBewertung.Berechne(
                belegung, input.Blocks, input.Slots,
                input.GewichtFrüheDoppel,
                input.GewichtSpäteDoppel,
                input.GewichtSpätePädEinheiten,
                input.StrafeHohlstunde,
                input.StrafeDoppelHohlstunde,
                input.StrafeDreifachHohlstunde,
                input.StrafeEinzelstunde,
                input.StrafeSpäteLkStunden,
                input.StrafeHauptfachSpät,
                input.HauptfachSpätAnteilProzent);
            return (b.Quality, b.BadUnits, belegung, "UNrPlan", input.Blocks);
        }

        // =====================================================
        // BUTTON 10 – FIX UNRN LÖSCHEN
        // =====================================================
        private void BtnFixUNrnLoeschen_Click(object sender, RoutedEventArgs e)
        {
            if (input == null)
            {
                MessageBox.Show("Bitte zuerst Excel-Datei laden (Button 2).");
                return;
            }

            // Alle Einzelklassen aus den Blöcken sammeln
            var alleEinzelKlassen = input.Blocks
                .SelectMany(b => b.Teile.SelectMany(t => t.Klassen))
                .Distinct()
                .OrderBy(k => k)
                .ToList();

            // Alle Klassenkombinationen (wie sie in der U-Verteilung vorkommen)
            var alleKombinationen = input.Blocks
                .SelectMany(b => b.Teile
                    .Where(t => t.Klassen.Count > 1)
                    .Select(t => string.Join(", ", t.Klassen.OrderBy(k => k))))
                .Distinct()
                .OrderBy(k => k)
                .ToList();

            // Kombinierte Liste: Einzelklassen + Kombinationen
            var alleKlassenOptionen = alleEinzelKlassen
                .Concat(alleKombinationen)
                .Distinct()
                .OrderBy(k => k)
                .ToList();

            // Alle Fächer sammeln
            var alleFächer = input.Blocks
                .SelectMany(b => b.Teile.Select(t => t.Fach.Trim()))
                .Where(f => !string.IsNullOrEmpty(f))
                .Distinct()
                .OrderBy(f => f)
                .ToList();

            // Dialog aufbauen
            var dlg = new Window
            {
                Title = "Fix UNrn löschen",
                Width = 420,
                Height = 500,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize
            };

            var stack = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };

            // Radiobuttons
            var rbAlles = new System.Windows.Controls.RadioButton
            {
                Content = "Alle Fix UNrn löschen (Spalten WTag/Stunde bleiben)",
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 8)
            };
            var rbSelektiv = new System.Windows.Controls.RadioButton
            {
                Content = "Nur bestimmte Klassen/Fächer löschen",
                Margin = new Thickness(0, 0, 0, 12)
            };

            // Klassen ListBox (Mehrfachauswahl)
            var lblKlassen = new System.Windows.Controls.TextBlock
            {
                Text = "Klassen / Kombinationen (Mehrfachauswahl möglich):",
                Margin = new Thickness(0, 0, 0, 4)
            };
            var lstKlassen = new System.Windows.Controls.ListBox
            {
                Height = 100,
                Margin = new Thickness(0, 0, 0, 8),
                SelectionMode = System.Windows.Controls.SelectionMode.Multiple
            };
            foreach (var k in alleKlassenOptionen)
                lstKlassen.Items.Add(k);

            // Fächer ListBox (Mehrfachauswahl)
            var lblFächer = new System.Windows.Controls.TextBlock
            {
                Text = "Fächer (Mehrfachauswahl möglich):",
                Margin = new Thickness(0, 0, 0, 4)
            };
            var lstFächer = new System.Windows.Controls.ListBox
            {
                Height = 100,
                Margin = new Thickness(0, 0, 0, 12),
                SelectionMode = System.Windows.Controls.SelectionMode.Multiple
            };
            foreach (var f in alleFächer)
                lstFächer.Items.Add(f);

            // Listen nur aktiv wenn selektiv
            rbSelektiv.Checked += (s, ev) => { lstKlassen.IsEnabled = true; lstFächer.IsEnabled = true; };
            rbAlles.Checked    += (s, ev) => { lstKlassen.IsEnabled = false; lstFächer.IsEnabled = false; };
            lstKlassen.IsEnabled = false;
            lstFächer.IsEnabled  = false;

            // Buttons
            var btnPanel = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var btnOk = new System.Windows.Controls.Button
            {
                Content = "Löschen",
                Width = 80,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var btnAbbrechen = new System.Windows.Controls.Button
            {
                Content = "Abbrechen",
                Width = 80
            };

            bool bestätigt = false;
            btnOk.Click        += (s, ev) => { bestätigt = true; dlg.Close(); };
            btnAbbrechen.Click += (s, ev) => dlg.Close();

            btnPanel.Children.Add(btnOk);
            btnPanel.Children.Add(btnAbbrechen);

            stack.Children.Add(rbAlles);
            stack.Children.Add(rbSelektiv);
            stack.Children.Add(lblKlassen);
            stack.Children.Add(lstKlassen);
            stack.Children.Add(lblFächer);
            stack.Children.Add(lstFächer);
            stack.Children.Add(btnPanel);
            dlg.Content = stack;
            dlg.ShowDialog();

            if (!bestätigt) return;

            bool allesLöschen = rbAlles.IsChecked == true;

            var gewählteKlassen = lstKlassen.SelectedItems.Cast<string>().ToHashSet();
            var gewählteFächer  = lstFächer.SelectedItems.Cast<string>()
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            try
            {
                using var wb = new XLWorkbook(excelPfad);

                if (!wb.Worksheets.Any(ws => ws.Name == "Fix UNrn"))
                {
                    MessageBox.Show("Tabelle 'Fix UNrn' nicht gefunden.");
                    return;
                }

                var sheet = wb.Worksheet("Fix UNrn");
                var letzteZeile = sheet.LastRowUsed()?.RowNumber() ?? 1;

                if (allesLöschen)
                {
                    // Nur UNr-Spalten (ab Spalte 3) löschen, WTag/Stunde behalten
                    for (int row = 2; row <= letzteZeile; row++)
                    {
                        var xlRow = sheet.Row(row);
                        int lastCol = xlRow.LastCellUsed()?.Address.ColumnNumber ?? 2;
                        for (int col = 3; col <= lastCol; col++)
                            xlRow.Cell(col).Clear();
                    }

                    wb.Save();
                    Log("Fix UNrn: alle UNrn gelöscht, WTag/Stunde-Spalten erhalten.");
                    TxtStatus.Text = "Alle Fix UNrn gelöscht.";
                    return;
                }

                // Selektiv: Lookup UNr → Block
                var unrZuBlock = new Dictionary<int, UnterrichtsBlock>();
                foreach (var block in input.Blocks)
                    unrZuBlock[block.UNr] = block;

                int gelöscht = 0;

                for (int row = 2; row <= letzteZeile; row++)
                {
                    var xlRow  = sheet.Row(row);
                    int lastCol = xlRow.LastCellUsed()?.Address.ColumnNumber ?? 2;

                    var verbleibende = new List<int>();

                    for (int col = 3; col <= lastCol; col++)
                    {
                        if (!int.TryParse(xlRow.Cell(col).GetString(), out int unr))
                            continue;

                        if (!unrZuBlock.TryGetValue(unr, out var block))
                        {
                            verbleibende.Add(unr);
                            continue;
                        }

                        bool klassenTreffer = false;
                        if (gewählteKlassen.Count > 0)
                        {
                            klassenTreffer = block.Teile.Any(t =>
                            {
                                string kombi = string.Join(", ", t.Klassen.OrderBy(k => k));
                                return gewählteKlassen.Contains(kombi) ||
                                       t.Klassen.Any(k => gewählteKlassen.Contains(k));
                            });
                        }

                        bool fachTreffer = gewählteFächer.Count > 0 &&
                            block.Teile.Any(t =>
                                gewählteFächer.Contains(t.Fach.Trim()));

                        // Löschen wenn Klasse ODER Fach passt
                        bool löschen = klassenTreffer || fachTreffer;

                        if (löschen)
                            gelöscht++;
                        else
                            verbleibende.Add(unr);
                    }

                    // Zeile neu schreiben
                    for (int col = 3; col <= lastCol; col++)
                        xlRow.Cell(col).Clear();

                    for (int i = 0; i < verbleibende.Count; i++)
                        xlRow.Cell(3 + i).Value = verbleibende[i];
                }

                wb.Save();

                string beschreibung = "";
                if (gewählteKlassen.Count > 0) beschreibung += $"Klassen: {string.Join(", ", gewählteKlassen)} ";
                if (gewählteFächer.Count > 0)  beschreibung += $"Fächer: {string.Join(", ", gewählteFächer)}";

                Log($"Fix UNrn gelöscht ({gelöscht} Einträge): {beschreibung.Trim()}");
                TxtStatus.Text = $"Fix UNrn gelöscht: {beschreibung.Trim()}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler:\n" + ex.Message);
            }
        }

        // =====================================================
        // ALTE SHEETS LÖSCHEN
        // =====================================================
        private void LöscheAlteSheets(string excelPfad, string prefix)
        {
            using var wb = new XLWorkbook(excelPfad);
            var zuLöschen = wb.Worksheets
                .Where(ws => ws.Name.StartsWith(prefix))
                .Select(ws => ws.Name)
                .ToList();

            foreach (var name in zuLöschen)
                wb.Worksheet(name).Delete();

            if (zuLöschen.Count > 0)
                wb.Save();
        }

        // =====================================================
        // BUTTON 7 – KLASSEN-UNTERRICHT ALS FIX SCHREIBEN
        // =====================================================
        private void BtnFixSchreiben_Click(object sender, RoutedEventArgs e)
        {
            if (letzteSolutions.Count == 0)
            {
                MessageBox.Show("Bitte zuerst Stundenplan erstellen (Button 3).");
                return;
            }

            // ── Lösung auswählen ──────────────────────────────
            var lösungsNamen = letzteSolutions.Select(s => s.label).ToList();
            string gewähltesLabel = ZeigeAuswahlDialog("Lösung wählen", lösungsNamen);
            if (gewähltesLabel == null) return;

            int lösungsIdx = lösungsNamen.IndexOf(gewähltesLabel);
            var gewählteLösung = letzteSolutions[lösungsIdx];

            // ── Klasse auswählen ──────────────────────────────
            var alleKlassen = gewählteLösung.blocks
                .SelectMany(b => b.Teile)
                .SelectMany(t => t.Klassen)
                .Distinct()
                .OrderBy(k => k)
                .ToList();

            string gewählteKlasse = ZeigeAuswahlDialog("Klasse wählen", alleKlassen);
            if (gewählteKlasse == null) return;

            // ── Fix-UNrn schreiben ────────────────────────────
            try
            {
                int geschrieben = SchreibeFixUnrn(
                    excelPfad,
                    gewählteLösung.belegung,
                    gewählteLösung.blocks,
                    input.Slots,
                    gewählteKlasse);

                Log($"Fix UNrn: {geschrieben} neue Einträge für Klasse {gewählteKlasse} aus [{gewählteLösung.label}] geschrieben.");
                TxtStatus.Text = $"Fix UNrn für {gewählteKlasse} geschrieben ({geschrieben} neue Slots).";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler:\n" + ex.Message);
            }
        }

        // Einfacher Auswahl-Dialog mit ListBox
        private string ZeigeAuswahlDialog(string titel, List<string> optionen)
        {
            var dlg = new Window
            {
                Title = titel,
                Width = 350,
                Height = 200,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize
            };

            var stack = new System.Windows.Controls.StackPanel { Margin = new Thickness(10) };
            var liste = new System.Windows.Controls.ListBox
            {
                ItemsSource = optionen,
                SelectedIndex = 0,
                Height = 120
            };
            var btn = new System.Windows.Controls.Button
            {
                Content = "OK",
                Margin = new Thickness(0, 8, 0, 0),
                Padding = new Thickness(20, 4, 20, 4),
                HorizontalAlignment = HorizontalAlignment.Right
            };

            string ergebnis = null;
            btn.Click += (s, e) => { ergebnis = liste.SelectedItem as string; dlg.Close(); };

            stack.Children.Add(liste);
            stack.Children.Add(btn);
            dlg.Content = stack;
            dlg.ShowDialog();

            return ergebnis;
        }

        // =====================================================
        // FIX-UNRN IN EXCEL SCHREIBEN
        // =====================================================
        private int SchreibeFixUnrn(
            string excelPfad,
            int[,] belegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            string klasse)
        {
            using var wb = new XLWorkbook(excelPfad);

            if (!wb.Worksheets.Any(ws => ws.Name == "Fix UNrn"))
                throw new Exception("Tabelle 'Fix UNrn' nicht gefunden.");

            var sheet = wb.Worksheet("Fix UNrn");

            // Bestehende Einträge einlesen
            var bestehend = new Dictionary<string, HashSet<int>>();
            var slotZeile = new Dictionary<string, int>();

            foreach (var row in sheet.RangeUsed()?.RowsUsed().Skip(1) ?? Enumerable.Empty<IXLRangeRow>())
            {
                string wtag = row.Cell(1).GetString().Trim();
                if (!int.TryParse(row.Cell(2).GetString(), out int std)) continue;

                string key = $"{wtag}_{std}";
                slotZeile[key] = row.RowNumber();

                var vorhandene = new HashSet<int>();
                int lastCol = row.LastCellUsed()?.Address.ColumnNumber ?? 2;
                for (int c = 3; c <= lastCol; c++)
                    if (int.TryParse(row.Cell(c).GetString(), out int u))
                        vorhandene.Add(u);

                bestehend[key] = vorhandene;
            }

            int geschrieben = 0;

            for (int s = 0; s < slots.Count; s++)
            {
                var slot = slots[s];
                string key = $"{slot.WTag}_{slot.Stunde}";

                // UNrn der gewählten Klasse in diesem Slot
                var unrnDieserKlasse = new List<int>();
                for (int b = 0; b < blocks.Count; b++)
                {
                    if (belegung[b, s] != 1) continue;
                    if (blocks[b].Teile.Any(t => t.Klassen.Contains(klasse)))
                        unrnDieserKlasse.Add(blocks[b].UNr);
                }

                if (unrnDieserKlasse.Count == 0) continue;

                // Zeile finden oder neu anlegen
                IXLRow xlRow;
                if (slotZeile.TryGetValue(key, out int zeilennr))
                {
                    xlRow = sheet.Row(zeilennr);
                }
                else
                {
                    int neueZeile = (sheet.RangeUsed()?.RowCount() ?? 1) + 1;
                    xlRow = sheet.Row(neueZeile);
                    xlRow.Cell(1).Value = slot.WTag;
                    xlRow.Cell(2).Value = slot.Stunde;
                    slotZeile[key] = neueZeile;
                    bestehend[key] = new HashSet<int>();
                }

                // Nur neue UNrn eintragen
                var vorhandene = bestehend[key];
                int nextCol = (xlRow.LastCellUsed()?.Address.ColumnNumber ?? 2) + 1;

                foreach (var unr in unrnDieserKlasse)
                {
                    if (vorhandene.Contains(unr)) continue;
                    xlRow.Cell(nextCol).Value = unr;
                    vorhandene.Add(unr);
                    nextCol++;
                    geschrieben++;
                }
            }

            wb.Save();
            return geschrieben;
        }

        // =====================================================
        // LÖSUNG IN ZEITSLOTS SCHREIBEN
        // =====================================================
        private void SetzeLoesungInSlots(int[,] belegung)
        {
            foreach (var slot in input.Slots)
                slot.BelegteUNrn.Clear();

            for (int b = 0; b < input.Blocks.Count; b++)
                for (int s = 0; s < input.Slots.Count; s++)
                    if (belegung[b, s] == 1)
                        input.Slots[s].BelegteUNrn.Add(input.Blocks[b].UNr);
        }
    }
}