using Google.OrTools.Sat;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Stundenplan_V2
{
    // =====================================================
    // EINSTELLUNGEN FÜR DIE VERBESSERUNG
    // =====================================================
    public class VerbesserungsOptionen
    {
        public VerbesserungsAlgorithmus Algorithmus { get; set; } = VerbesserungsAlgorithmus.HillClimbing;
        public VerbesserungsZiel Ziel { get; set; } = VerbesserungsZiel.Gesamt;
        public int ZeitlimitSekunden { get; set; } = 30;

        // Einschränkungen
        public HashSet<string> NurLehrer { get; set; } = new(); // leer = alle
        public HashSet<string> NurKlassen { get; set; } = new(); // leer = alle
        public bool FixUNrnRespektieren { get; set; } = true;

        // Simulated Annealing
        public double StartTemperatur { get; set; } = 100.0;
        public double Abkühlrate { get; set; } = 0.995;

        // LNS
        public int LnsZeitlimitSekunden { get; set; } = 10;
        public int LnsNachbarschaftsGröße { get; set; } = 10; // Anzahl freizugebender Blöcke
    }

    public enum VerbesserungsAlgorithmus
    {
        HillClimbing,
        SimulatedAnnealing,
        LargeNeighborhoodSearch
    }

    public enum VerbesserungsZiel
    {
        Gesamt,
        Hohlstunden,
        SpäteDoppelstunden,
        SpätePädEinheiten,
        Einzelstunden,
        HauptfachSpät
    }

    // =====================================================
    // ERGEBNIS DER VERBESSERUNG
    // =====================================================
    public class VerbesserungsErgebnis
    {
        public int[,] BesteBelegung { get; set; }
        public int AusgangsQualität { get; set; }
        public int EndQualität { get; set; }
        public int Verbesserung => EndQualität - AusgangsQualität;
        public int Iterationen { get; set; }
        public int AkzeptierteVerbesserungen { get; set; }
        public List<string> Log { get; set; } = new();
    }

    // =====================================================
    // HAUPT-KLASSE
    // =====================================================
    public static class PlanVerbesserung
    {
        private static Random _rnd = new Random();

        public static VerbesserungsErgebnis Verbessere(
            int[,] ausgangsBelegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            StundenplanInput input,
            VerbesserungsOptionen optionen,
            Action<string> log)
        {
            var ergebnis = new VerbesserungsErgebnis();
            int B = blocks.Count;
            int S = slots.Count;

            // Ausgangsbelegung kopieren
            var belegung = KopiereBelegung(ausgangsBelegung, B, S);

            // Ausgangsqualität berechnen
            ergebnis.AusgangsQualität = BerechneZiel(belegung, blocks, slots, input, optionen.Ziel);
            ergebnis.BesteBelegung   = KopiereBelegung(belegung, B, S);

            log($"Ausgangsqualität: {ergebnis.AusgangsQualität}");
            log($"Algorithmus: {optionen.Algorithmus}, Ziel: {optionen.Ziel}");

            // Fix-Slots ermitteln
            var fixSlots = new HashSet<(int b, int s)>();
            if (optionen.FixUNrnRespektieren)
            {
                for (int s = 0; s < S; s++)
                    foreach (var unr in slots[s].FixUNrn)
                        for (int b = 0; b < B; b++)
                            if (blocks[b].UNr == unr)
                                fixSlots.Add((b, s));
            }

            // Erlaubte Blöcke ermitteln
            var erlaubteBlöcke = ErmittleErlaubteBlöcke(blocks, optionen);

            log($"Erlaubte Blöcke: {erlaubteBlöcke.Count} von {B}");

            switch (optionen.Algorithmus)
            {
                case VerbesserungsAlgorithmus.HillClimbing:
                    HillClimbing(belegung, ergebnis, blocks, slots, input, optionen,
                        erlaubteBlöcke, fixSlots, B, S, log);
                    break;

                case VerbesserungsAlgorithmus.SimulatedAnnealing:
                    SimulatedAnnealing(belegung, ergebnis, blocks, slots, input, optionen,
                        erlaubteBlöcke, fixSlots, B, S, log);
                    break;

                case VerbesserungsAlgorithmus.LargeNeighborhoodSearch:
                    LargeNeighborhoodSearch(belegung, ergebnis, blocks, slots, input, optionen,
                        erlaubteBlöcke, fixSlots, B, S, log);
                    break;
            }

            log($"Endqualität: {ergebnis.EndQualität} " +
                $"(Verbesserung: {ergebnis.Verbesserung:+0;-0;0})");
            log($"Iterationen: {ergebnis.Iterationen}, " +
                $"Akzeptierte Verbesserungen: {ergebnis.AkzeptierteVerbesserungen}");

            return ergebnis;
        }

        // =====================================================
        // HILL CLIMBING
        // =====================================================
        private static void HillClimbing(
            int[,] belegung,
            VerbesserungsErgebnis ergebnis,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            StundenplanInput input,
            VerbesserungsOptionen optionen,
            List<int> erlaubteBlöcke,
            HashSet<(int, int)> fixSlots,
            int B, int S,
            Action<string> log)
        {
            var deadline = DateTime.Now.AddSeconds(optionen.ZeitlimitSekunden);
            int bestQuality = ergebnis.AusgangsQualität;
            bool verbesserungGefunden = true;

            while (verbesserungGefunden && DateTime.Now < deadline)
            {
                verbesserungGefunden = false;
                ergebnis.Iterationen++;

                // Alle möglichen Tausche durchprobieren
                var tausche = ErzeugeTausche(belegung, erlaubteBlöcke, fixSlots, B, S);

                foreach (var (b1, s1, b2, s2) in tausche)
                {
                    if (DateTime.Now >= deadline) break;

                    // Tausch durchführen
                    FühreTauschDurch(belegung, b1, s1, b2, s2);

                    // Prüfen ob gültig und besser
                    if (IstGültig(belegung, blocks, slots, input, fixSlots, B, S))
                    {
                        int newQuality = BerechneZiel(belegung, blocks, slots, input, optionen.Ziel);

                        if (newQuality > bestQuality)
                        {
                            bestQuality = newQuality;
                            ergebnis.BesteBelegung = KopiereBelegung(belegung, B, S);
                            ergebnis.AkzeptierteVerbesserungen++;
                            verbesserungGefunden = true;
                            log($"  Verbesserung gefunden: {newQuality}");
                            break; // Neustart mit verbessertem Plan
                        }
                    }

                    // Tausch rückgängig machen
                    FühreTauschDurch(belegung, b1, s1, b2, s2);
                }
            }

            ergebnis.EndQualität = bestQuality;
        }

        // =====================================================
        // SIMULATED ANNEALING
        // =====================================================
        private static void SimulatedAnnealing(
            int[,] belegung,
            VerbesserungsErgebnis ergebnis,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            StundenplanInput input,
            VerbesserungsOptionen optionen,
            List<int> erlaubteBlöcke,
            HashSet<(int, int)> fixSlots,
            int B, int S,
            Action<string> log)
        {
            var deadline = DateTime.Now.AddSeconds(optionen.ZeitlimitSekunden);
            double temperatur = optionen.StartTemperatur;
            int aktuelleQualität = ergebnis.AusgangsQualität;
            int besteQualität    = aktuelleQualität;

            while (DateTime.Now < deadline)
            {
                ergebnis.Iterationen++;

                // Zufälligen Tausch wählen
                var tausch = WähleZufälligenTausch(belegung, erlaubteBlöcke, fixSlots, B, S);
                if (tausch == null) break;

                var (b1, s1, b2, s2) = tausch.Value;

                FühreTauschDurch(belegung, b1, s1, b2, s2);

                if (IstGültig(belegung, blocks, slots, input, fixSlots, B, S))
                {
                    int neueQualität = BerechneZiel(belegung, blocks, slots, input, optionen.Ziel);
                    int delta = neueQualität - aktuelleQualität;

                    // Akzeptanzkriterium
                    bool akzeptieren = delta > 0 ||
                        _rnd.NextDouble() < Math.Exp(delta / temperatur);

                    if (akzeptieren)
                    {
                        aktuelleQualität = neueQualität;
                        ergebnis.AkzeptierteVerbesserungen++;

                        if (neueQualität > besteQualität)
                        {
                            besteQualität = neueQualität;
                            ergebnis.BesteBelegung = KopiereBelegung(belegung, B, S);
                            log($"  Neue beste Qualität: {besteQualität} (T={temperatur:F1})");
                        }
                    }
                    else
                    {
                        FühreTauschDurch(belegung, b1, s1, b2, s2);
                    }
                }
                else
                {
                    FühreTauschDurch(belegung, b1, s1, b2, s2);
                }

                temperatur *= optionen.Abkühlrate;
                if (temperatur < 0.01) temperatur = 0.01;
            }

            ergebnis.EndQualität = besteQualität;
        }

        // =====================================================
        // LARGE NEIGHBORHOOD SEARCH
        // =====================================================
        private static void LargeNeighborhoodSearch(
            int[,] belegung,
            VerbesserungsErgebnis ergebnis,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            StundenplanInput input,
            VerbesserungsOptionen optionen,
            List<int> erlaubteBlöcke,
            HashSet<(int, int)> fixSlots,
            int B, int S,
            Action<string> log)
        {
            var deadline = DateTime.Now.AddSeconds(optionen.ZeitlimitSekunden);
            int besteQualität = ergebnis.AusgangsQualität;
            ergebnis.BesteBelegung = KopiereBelegung(belegung, B, S);

            while (DateTime.Now < deadline)
            {
                ergebnis.Iterationen++;

                // Zufällige Nachbarschaft wählen: n Blöcke freigeben
                int n = Math.Min(optionen.LnsNachbarschaftsGröße, erlaubteBlöcke.Count);
                var freigegebeneBlöcke = erlaubteBlöcke
                    .OrderBy(_ => _rnd.Next())
                    .Take(n)
                    .ToHashSet();

                // OR-Tools Sub-Solver mit fixierten Blöcken
                var neueBelegung = LnsSubSolver(
                    ergebnis.BesteBelegung,
                    blocks, slots, input,
                    freigegebeneBlöcke, fixSlots,
                    optionen.LnsZeitlimitSekunden,
                    B, S);

                if (neueBelegung == null)
                {
                    log($"  Iteration {ergebnis.Iterationen}: keine Lösung gefunden");
                    continue;
                }

                int neueQualität = BerechneZiel(neueBelegung, blocks, slots, input, optionen.Ziel);

                if (neueQualität > besteQualität)
                {
                    besteQualität = neueQualität;
                    ergebnis.BesteBelegung = neueBelegung;
                    ergebnis.AkzeptierteVerbesserungen++;
                    log($"  Iteration {ergebnis.Iterationen}: Verbesserung auf {besteQualität}");

                    // Aktuelle Belegung auf beste setzen
                    for (int b = 0; b < B; b++)
                        for (int s = 0; s < S; s++)
                            belegung[b, s] = ergebnis.BesteBelegung[b, s];
                }
            }

            ergebnis.EndQualität = besteQualität;
        }

        // =====================================================
        // LNS SUB-SOLVER
        // =====================================================
        private static int[,] LnsSubSolver(
            int[,] ausgangsBelegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            StundenplanInput input,
            HashSet<int> freigegebeneBlöcke,
            HashSet<(int, int)> fixSlots,
            int zeitlimit,
            int B, int S)
        {
            var model = new CpModel();

            BoolVar[,] x = new BoolVar[B, S];
            for (int b = 0; b < B; b++)
                for (int s = 0; s < S; s++)
                    x[b, s] = model.NewBoolVar($"x_{b}_{s}");

            // Wochenstunden
            for (int b = 0; b < B; b++)
                model.Add(LinearExpr.Sum(
                    Enumerable.Range(0, S).Select(s => x[b, s])) == blocks[b].Wst);

            // Nicht freigegebene Blöcke fixieren
            for (int b = 0; b < B; b++)
            {
                if (freigegebeneBlöcke.Contains(b)) continue;
                for (int s = 0; s < S; s++)
                    model.Add(x[b, s] == ausgangsBelegung[b, s]);
            }

            // Fix-Slots
            foreach (var (fb, fs) in fixSlots)
                model.Add(x[fb, fs] == 1);

            // Lehrerregel
            for (int s = 0; s < S; s++)
            {
                var map = new Dictionary<string, List<int>>();
                for (int b = 0; b < B; b++)
                    foreach (var l in blocks[b].Teile.Select(t => t.Lehrer).Distinct())
                    {
                        if (!map.ContainsKey(l)) map[l] = new List<int>();
                        map[l].Add(b);
                    }
                foreach (var kv in map)
                    model.Add(LinearExpr.Sum(kv.Value.Select(b => x[b, s])) <= 1);
            }

            // Klassenregel
            ClassConstraint.Add(model, x, blocks, S);

            // Sperrslots
            TimeConstraint.AddBlockedSlots(model, x, blocks, slots, B, S);

            // Tagesregel
            var tage = slots.Select(z => z.WTag).Distinct();
            foreach (var tag in tage)
            {
                var daySlots = slots
                    .Select((z, i) => new { z, i })
                    .Where(z => z.z.WTag == tag)
                    .Select(z => z.i).ToList();

                for (int b = 0; b < B; b++)
                {
                    int maxD = blocks[b].Teile.Max(t => t.MaxDoppel);
                    int limit = (maxD == 0 && blocks[b].Wst >= 2) ? 1 : 2;
                    model.Add(LinearExpr.Sum(daySlots.Select(s => x[b, s])) <= limit);
                }
            }

            // Zielfunktion – Hohlstunden und Doppelstunden minimieren
            var earlyVars = new List<BoolVar>();
            var lateVars  = new List<BoolVar>();

            var dVars = new BoolVar[B, S];
            for (int b = 0; b < B; b++)
            {
                for (int s = 0; s < S - 1; s++)
                {
                    if (slots[s].WTag == slots[s + 1].WTag &&
                        slots[s].Stunde + 1 == slots[s + 1].Stunde)
                    {
                        dVars[b, s] = model.NewBoolVar($"d_{b}_{s}");
                        model.Add(x[b, s] == 1).OnlyEnforceIf(dVars[b, s]);
                        model.Add(x[b, s + 1] == 1).OnlyEnforceIf(dVars[b, s]);
                        model.Add(x[b, s] + x[b, s + 1] - dVars[b, s] <= 1);

                        if (slots[s].Stunde <= 5) earlyVars.Add(dVars[b, s]);
                        else lateVars.Add(dVars[b, s]);
                    }
                }
            }

            var qualExpr = LinearExpr.Sum(earlyVars)
                - LinearExpr.Sum(lateVars) * input.GewichtSpäteDoppel;
            model.Maximize(qualExpr);

            var solver = new CpSolver();
            solver.StringParameters =
                $"max_time_in_seconds:{zeitlimit} num_search_workers:4";

            var status = solver.Solve(model);
            if (status != CpSolverStatus.Optimal && status != CpSolverStatus.Feasible)
                return null;

            var result = new int[B, S];
            for (int b = 0; b < B; b++)
                for (int s = 0; s < S; s++)
                    result[b, s] = (int)solver.Value(x[b, s]);

            return result;
        }

        // =====================================================
        // HILFSMETHODEN
        // =====================================================

        private static List<int> ErmittleErlaubteBlöcke(
            List<UnterrichtsBlock> blocks,
            VerbesserungsOptionen optionen)
        {
            var result = new List<int>();
            for (int b = 0; b < blocks.Count; b++)
            {
                var block = blocks[b];

                // Lehrer-Filter
                if (optionen.NurLehrer.Count > 0 &&
                    !block.Teile.Any(t => optionen.NurLehrer.Contains(t.Lehrer)))
                    continue;

                // Klassen-Filter
                if (optionen.NurKlassen.Count > 0 &&
                    !block.Teile.Any(t =>
                        t.Klassen.Any(k => optionen.NurKlassen.Contains(k))))
                    continue;

                result.Add(b);
            }
            return result;
        }

        private static List<(int b1, int s1, int b2, int s2)> ErzeugeTausche(
            int[,] belegung,
            List<int> erlaubteBlöcke,
            HashSet<(int, int)> fixSlots,
            int B, int S)
        {
            var tausche = new List<(int, int, int, int)>();

            foreach (int b1 in erlaubteBlöcke)
            {
                for (int s1 = 0; s1 < S; s1++)
                {
                    if (belegung[b1, s1] != 1) continue;
                    if (fixSlots.Contains((b1, s1))) continue;

                    foreach (int b2 in erlaubteBlöcke)
                    {
                        for (int s2 = s1 + 1; s2 < S; s2++)
                        {
                            if (belegung[b2, s2] != 1) continue;
                            if (fixSlots.Contains((b2, s2))) continue;
                            if (b1 == b2 && s1 == s2) continue;

                            tausche.Add((b1, s1, b2, s2));
                        }
                    }
                }
            }

            // Zufällig mischen für Diversität
            return tausche.OrderBy(_ => _rnd.Next()).ToList();
        }

        private static (int b1, int s1, int b2, int s2)? WähleZufälligenTausch(
            int[,] belegung,
            List<int> erlaubteBlöcke,
            HashSet<(int, int)> fixSlots,
            int B, int S)
        {
            // Alle belegten Slots der erlaubten Blöcke sammeln
            var belegteSlots = new List<(int b, int s)>();
            foreach (int b in erlaubteBlöcke)
                for (int s = 0; s < S; s++)
                    if (belegung[b, s] == 1 && !fixSlots.Contains((b, s)))
                        belegteSlots.Add((b, s));

            if (belegteSlots.Count < 2) return null;

            int idx1 = _rnd.Next(belegteSlots.Count);
            int idx2 = _rnd.Next(belegteSlots.Count - 1);
            if (idx2 >= idx1) idx2++;

            var (b1, s1) = belegteSlots[idx1];
            var (b2, s2) = belegteSlots[idx2];

            return (b1, s1, b2, s2);
        }

        private static void FühreTauschDurch(int[,] belegung, int b1, int s1, int b2, int s2)
        {
            int tmp = belegung[b1, s1];
            belegung[b1, s1] = belegung[b2, s2];
            belegung[b2, s2] = tmp;
        }

        private static bool IstGültig(
            int[,] belegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            StundenplanInput input,
            HashSet<(int, int)> fixSlots,
            int B, int S)
        {
            // Lehrerregel
            for (int s = 0; s < S; s++)
            {
                var lehrer = new Dictionary<string, int>();
                for (int b = 0; b < B; b++)
                {
                    if (belegung[b, s] != 1) continue;
                    foreach (var t in blocks[b].Teile)
                    {
                        if (lehrer.ContainsKey(t.Lehrer)) return false;
                        lehrer[t.Lehrer] = b;
                    }
                }
            }

            // Klassenregel
            for (int s = 0; s < S; s++)
            {
                var klassen = new Dictionary<string, int>();
                for (int b = 0; b < B; b++)
                {
                    if (belegung[b, s] != 1) continue;
                    foreach (var t in blocks[b].Teile)
                        foreach (var k in t.Klassen)
                        {
                            if (klassen.ContainsKey(k) && klassen[k] != b)
                                return false;
                            klassen[k] = b;
                        }
                }
            }

            // Sperrslots
            for (int b = 0; b < B; b++)
                for (int s = 0; s < S; s++)
                {
                    if (belegung[b, s] != 1) continue;
                    foreach (var t in blocks[b].Teile)
                    {
                        if (slots[s].LehrerWunsch.TryGetValue(t.Lehrer, out int lw) && lw == -3)
                            return false;
                        foreach (var k in t.Klassen)
                            if (slots[s].KlassenWunsch.TryGetValue(k, out int kw) && kw == -3)
                                return false;
                    }
                }

            // Fix-Slots
            foreach (var (fb, fs) in fixSlots)
                if (belegung[fb, fs] != 1) return false;

            // Tagesregel
            var tage = slots.Select(z => z.WTag).Distinct();
            foreach (var tag in tage)
            {
                var daySlots = Enumerable.Range(0, S)
                    .Where(s => slots[s].WTag == tag).ToList();

                for (int b = 0; b < B; b++)
                {
                    int maxD = blocks[b].Teile.Max(t => t.MaxDoppel);
                    int limit = (maxD == 0 && blocks[b].Wst >= 2) ? 1 : 2;
                    if (daySlots.Sum(s => belegung[b, s]) > limit)
                        return false;
                }
            }

            // Große Pausen
            if (input.GrossePausen != null)
            {
                for (int b = 0; b < B; b++)
                {
                    if (blocks[b].DoppelÜberPauseErlaubt) continue;
                    for (int s = 0; s < S - 1; s++)
                    {
                        if (belegung[b, s] != 1 || belegung[b, s + 1] != 1) continue;
                        if (slots[s].WTag != slots[s + 1].WTag) continue;
                        bool istPause = input.GrossePausen.Any(p =>
                            p.stundeVor == slots[s].Stunde &&
                            p.stundeNach == slots[s + 1].Stunde);
                        if (istPause) return false;
                    }
                }
            }

            return true;
        }

        private static int BerechneZiel(
            int[,] belegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            StundenplanInput input,
            VerbesserungsZiel ziel)
        {
            switch (ziel)
            {
                case VerbesserungsZiel.Gesamt:
                    return PlanBewertung.Berechne(belegung, blocks, slots,
                        input.GewichtFrüheDoppel,
                        input.GewichtSpäteDoppel,
                        input.GewichtSpätePädEinheiten,
                        input.StrafeHohlstunde,
                        input.StrafeDoppelHohlstunde,
                        input.StrafeDreifachHohlstunde,
                        input.StrafeEinzelstunde,
                        input.StrafeSpäteLkStunden,
                        input.StrafeHauptfachSpät,
                        input.HauptfachSpätAnteilProzent).Quality;

                case VerbesserungsZiel.SpäteDoppelstunden:
                    return -PlanBewertung.Berechne(belegung, blocks, slots, 1, 1, 0).Late;

                case VerbesserungsZiel.SpätePädEinheiten:
                    return -PlanBewertung.Berechne(belegung, blocks, slots, 0, 0, 1).BadUnits;

                case VerbesserungsZiel.Hohlstunden:
                    return -BerechnHohlstunden(belegung, blocks, slots);

                case VerbesserungsZiel.Einzelstunden:
                    return -BerechneEinzelstunden(belegung, blocks, slots);

                case VerbesserungsZiel.HauptfachSpät:
                    return -BerechneHauptfachSpät(belegung, blocks, slots,
                        input.HauptfachSpätAnteilProzent);

                default:
                    return PlanBewertung.Berechne(belegung, blocks, slots,
                        input.GewichtFrüheDoppel,
                        input.GewichtSpäteDoppel,
                        input.GewichtSpätePädEinheiten).Quality;
            }
        }

        private static int BerechnHohlstunden(
            int[,] belegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots)
        {
            int gesamt = 0;
            int B = blocks.Count;
            int S = slots.Count;

            var alleLehrer = blocks.SelectMany(b => b.Teile.Select(t => t.Lehrer))
                .Distinct().ToList();
            var tage = slots.Select(s => s.WTag).Distinct().ToList();

            foreach (var lehrer in alleLehrer)
            {
                var lehrerBlöcke = Enumerable.Range(0, B)
                    .Where(b => blocks[b].Teile.Any(t => t.Lehrer == lehrer))
                    .ToList();

                foreach (var tag in tage)
                {
                    var tagesSlots = Enumerable.Range(0, S)
                        .Where(s => slots[s].WTag == tag)
                        .OrderBy(s => slots[s].Stunde)
                        .ToList();

                    var mitUnterricht = new HashSet<int>();
                    foreach (var s in tagesSlots)
                        foreach (var b in lehrerBlöcke)
                            if (belegung[b, s] == 1)
                                mitUnterricht.Add(slots[s].Stunde);

                    if (mitUnterricht.Count < 2) continue;

                    int ersteStd = mitUnterricht.Min();
                    int letzteStd = mitUnterricht.Max();

                    for (int std = ersteStd + 1; std < letzteStd; std++)
                        if (!mitUnterricht.Contains(std))
                            gesamt++;
                }
            }

            return gesamt;
        }

        private static int BerechneEinzelstunden(
            int[,] belegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots)
        {
            int gesamt = 0;
            int B = blocks.Count;
            int S = slots.Count;

            var alleLehrer = blocks.SelectMany(b => b.Teile.Select(t => t.Lehrer))
                .Distinct().ToList();
            var tage = slots.Select(s => s.WTag).Distinct().ToList();

            foreach (var lehrer in alleLehrer)
            {
                var lehrerBlöcke = Enumerable.Range(0, B)
                    .Where(b => blocks[b].Teile.Any(t => t.Lehrer == lehrer))
                    .ToList();

                foreach (var tag in tage)
                {
                    var tagesSlots = Enumerable.Range(0, S)
                        .Where(s => slots[s].WTag == tag).ToList();

                    int anzahl = tagesSlots.Sum(s =>
                        lehrerBlöcke.Any(b => belegung[b, s] == 1) ? 1 : 0);

                    if (anzahl == 1) gesamt++;
                }
            }

            return gesamt;
        }

        private static int BerechneHauptfachSpät(
            int[,] belegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            int anteilProzent)
        {
            int gesamt = 0;
            int B = blocks.Count;
            int S = slots.Count;
            var hauptfächer = new HashSet<string> { "D", "E", "M", "F" };

            // Pädagogische Einheiten Typ 2 aufbauen
            var einheiten = new Dictionary<(string klasse, string fach), List<int>>();
            for (int b = 0; b < B; b++)
                foreach (var t in blocks[b].Teile)
                {
                    string fach = t.Fach.Trim();
                    if (!hauptfächer.Contains(fach)) continue;
                    foreach (var klasse in t.Klassen)
                    {
                        var key = (klasse, fach);
                        if (!einheiten.ContainsKey(key)) einheiten[key] = new List<int>();
                        if (!einheiten[key].Contains(b)) einheiten[key].Add(b);
                    }
                }

            foreach (var kv in einheiten)
            {
                var blockIds = kv.Value;
                int gesamtWst = blockIds.Sum(b => blocks[b].Wst);
                if (gesamtWst == 0) continue;

                int erlaubtSpät = (int)Math.Floor(gesamtWst * anteilProzent / 100.0);

                int spätStunden = 0;
                foreach (int b in blockIds)
                    for (int s = 0; s < S; s++)
                        if (belegung[b, s] == 1 && slots[s].Stunde >= 5)
                            spätStunden++;

                int überschuss = Math.Max(0, spätStunden - erlaubtSpät);
                gesamt += überschuss;
            }

            return gesamt;
        }

        private static int[,] KopiereBelegung(int[,] quelle, int B, int S)
        {
            var kopie = new int[B, S];
            for (int b = 0; b < B; b++)
                for (int s = 0; s < S; s++)
                    kopie[b, s] = quelle[b, s];
            return kopie;
        }
    }
}
