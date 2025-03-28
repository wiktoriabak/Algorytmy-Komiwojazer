using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        string filePath = "C:\\Users\\wiki2\\OneDrive\\Pulpit\\Dane_TSP_48.xlsx";
        double[,] distances = LoadDistancesFromExcel(filePath);

        // Parametry algorytmu
        int numCities = distances.GetLength(0);
        string[] neighborTypes = { "swap", "reverse", "insert" };
        int[] maxIterationsArray = { 1000, 5000, 10000, 20000 };
        int[] maxNoImprovementArray = { 500, 2500, 5000, 10000 };
        int[] numRestartsArray = { 20, 50, 100, 500 };
        int rep = 1;

        string summaryOutputPath = "C:\\Users\\wiki2\\OneDrive\\Pulpit\\Wyniki_TSP.xlsx";
        string detailedOutputPath = "C:\\Users\\wiki2\\OneDrive\\Pulpit\\Wyniki_TSP_Detale.xlsx";

        bool summaryFileExists = File.Exists(summaryOutputPath);
        bool detailedFileExists = File.Exists(detailedOutputPath);

        int totalCombinations = neighborTypes.Length * maxIterationsArray.Length * maxNoImprovementArray.Length * numRestartsArray.Length * rep;
        int completedRuns = 0;

        using (var summaryPackage = new ExcelPackage(new FileInfo(summaryOutputPath)))
        {
            var summaryWorksheet = summaryPackage.Workbook.Worksheets.Count > 0
                ? summaryPackage.Workbook.Worksheets[0]
                : summaryPackage.Workbook.Worksheets.Add("Wyniki");

            if (!summaryFileExists)
            {
                summaryWorksheet.Cells[1, 1].Value = "Data";
                summaryWorksheet.Cells[1, 2].Value = "Najlepszy Koszt";
                summaryWorksheet.Cells[1, 3].Value = "Czas (ms)";
                summaryWorksheet.Cells[1, 4].Value = "Rodzaj Sąsiedztwa";
                summaryWorksheet.Cells[1, 5].Value = "Liczba Restartów";
                summaryWorksheet.Cells[1, 6].Value = "Maks. Iteracje Bez Poprawy";
                summaryWorksheet.Cells[1, 7].Value = "Maksymalna liczba Iteracji";
                summaryWorksheet.Cells[1, 8].Value = "Liczba Miast";
                summaryWorksheet.Cells[1, 9].Value = "Kolejność Miast (Najlepszy Wynik)";
            }

            using (var detailedPackage = new ExcelPackage(new FileInfo(detailedOutputPath)))
            {
                var detailedWorksheet = detailedPackage.Workbook.Worksheets.Count > 0
                    ? detailedPackage.Workbook.Worksheets[0]
                    : detailedPackage.Workbook.Worksheets.Add("Detale");

                if (!detailedFileExists)
                {
                    detailedWorksheet.Cells[1, 1].Value = "Nr Uruchomienia";
                    detailedWorksheet.Cells[1, 2].Value = "Koszt";
                    detailedWorksheet.Cells[1, 3].Value = "Iteracje";
                    detailedWorksheet.Cells[1, 4].Value = "Kolejność Miast";
                    detailedWorksheet.Cells[1, 5].Value = "Liczba Miast";
                    detailedWorksheet.Cells[1, 6].Value = "Maks. Iteracje Bez Poprawy";
                    detailedWorksheet.Cells[1, 7].Value = "Maks. Iteracje";
                    detailedWorksheet.Cells[1, 8].Value = "Rodzaj Sąsiedztwa";
                    detailedWorksheet.Cells[1, 9].Value = "Liczba Restartów";
                }

                int detailRow = detailedWorksheet.Dimension?.Rows + 1 ?? 2;

                foreach (string neighborType in neighborTypes)
                {
                    foreach (int maxIterations in maxIterationsArray)
                    {
                        foreach (int maxNoImprovement in maxNoImprovementArray)
                        {
                            foreach (int numRestarts in numRestartsArray)
                            {
                                double bestCost = double.MaxValue;
                                List<int> bestRoute = null;
                                Stopwatch stopwatch = Stopwatch.StartNew();

                                for (int i = 0; i < rep; i++)
                                {
                                    (List<int> route, double cost, int iterationCount) = MultistartHillClimbing(distances, numCities, maxNoImprovement, maxIterations, neighborType, numRestarts);

                                    if (cost < bestCost)
                                    {
                                        bestCost = cost;
                                        bestRoute = new List<int>(route);
                                    }

                                    detailedWorksheet.Cells[detailRow, 1].Value = i + 1;
                                    detailedWorksheet.Cells[detailRow, 2].Value = cost;
                                    detailedWorksheet.Cells[detailRow, 3].Value = iterationCount;
                                    detailedWorksheet.Cells[detailRow, 4].Value = string.Join(",", route);
                                    detailedWorksheet.Cells[detailRow, 5].Value = numCities;
                                    detailedWorksheet.Cells[detailRow, 6].Value = maxNoImprovement;
                                    detailedWorksheet.Cells[detailRow, 7].Value = maxIterations;
                                    detailedWorksheet.Cells[detailRow, 8].Value = neighborType;
                                    detailedWorksheet.Cells[detailRow, 9].Value = numRestarts;
                                    detailRow++;
                                    completedRuns++;
                                    double progress = (double)completedRuns / totalCombinations * 100;
                                    Console.SetCursorPosition(0, Console.CursorTop);
                                    Console.Write($"Postęp: {progress:F2}%");
                                }

                                stopwatch.Stop();

                                int summaryRow = summaryWorksheet.Dimension?.Rows + 1 ?? 2;
                                summaryWorksheet.Cells[summaryRow, 1].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                summaryWorksheet.Cells[summaryRow, 2].Value = bestCost;
                                summaryWorksheet.Cells[summaryRow, 3].Value = stopwatch.ElapsedMilliseconds;
                                summaryWorksheet.Cells[summaryRow, 4].Value = neighborType;
                                summaryWorksheet.Cells[summaryRow, 5].Value = numRestarts;
                                summaryWorksheet.Cells[summaryRow, 6].Value = maxNoImprovement;
                                summaryWorksheet.Cells[summaryRow, 7].Value = maxIterations;
                                summaryWorksheet.Cells[summaryRow, 8].Value = numCities;
                                summaryWorksheet.Cells[summaryRow, 9].Value = string.Join(",", bestRoute);

                                detailedPackage.Save();
                                summaryPackage.Save();
                            }
                        }
                    }
                }
            }
        }

        Console.WriteLine("\nZakończono zapisywanie wyników.");
    }

    static double[,] LoadDistancesFromExcel(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            if (package.Workbook.Worksheets.Count == 0)
            {
                throw new Exception("Plik Excel nie zawiera żadnych arkuszy.");
            }

            var worksheet = package.Workbook.Worksheets[0];
            int rows = worksheet.Dimension.Rows - 1;
            int cols = worksheet.Dimension.Columns - 1;

            if (rows <= 0 || cols <= 0)
            {
                throw new Exception("Arkusz jest pusty lub nieprawidłowy.");
            }

            double[,] distances = new double[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    var cellValue = worksheet.Cells[i + 2, j + 2].Value;
                    if (cellValue == null || !double.TryParse(cellValue.ToString(), out double distance))
                    {
                        throw new Exception($"Nieprawidłowa wartość w komórce ({i + 2}, {j + 2}).");
                    }
                    distances[i, j] = distance;
                }
            }

            return distances;
        }
    }

    static (List<int>, double, int) MultistartHillClimbing(double[,] distances, int numCities, int maxNoImprovement, int maxIterations, string neighborType, int numRestarts)
    {
        double bestCost = double.MaxValue;
        List<int> bestRoute = null;
        int totalIterations = 0;

        for (int restart = 0; restart < numRestarts; restart++)
        {
            List<int> currentRoute = GenerateRandomRoute(numCities);
            double currentCost = CalculateRouteCost(currentRoute, distances);

            int noImprovementCounter = 0;
            int iterationCount = 0;

            while (noImprovementCounter < maxNoImprovement && iterationCount < maxIterations)
            {
                iterationCount++;
                var (newRoute, newCost) = EvaluateRandomNeighbor(currentRoute, distances, neighborType);

                if (newCost < currentCost)
                {
                    currentRoute = newRoute;
                    currentCost = newCost;
                    noImprovementCounter = 0;
                }
                else
                {
                    noImprovementCounter++;
                }
            }
            if (currentCost < bestCost)
            {
                bestCost = currentCost;
                bestRoute = new List<int>(currentRoute);
            }

            totalIterations += iterationCount;
        }

        return (bestRoute, bestCost, totalIterations);
    }


    static List<int> GenerateRandomRoute(int numCities)
    {
        List<int> route = new List<int>();
        for (int i = 0; i < numCities; i++)
            route.Add(i);

        Random random = new Random();
        for (int i = 0; i < route.Count; i++)
        {
            int swapIndex = random.Next(route.Count);
            (route[i], route[swapIndex]) = (route[swapIndex], route[i]);
        }

        return route;
    }

    static double CalculateRouteCost(List<int> route, double[,] distances)
    {
        double cost = 0;
        for (int i = 0; i < route.Count - 1; i++)
        {
            cost += distances[route[i], route[i + 1]];
        }
        cost += distances[route[route.Count - 1], route[0]];
        return cost;
    }

    static (List<int>, double) EvaluateRandomNeighbor(List<int> route, double[,] distances, string neighborType)
    {
        Random random = new Random();
        int i = random.Next(0, route.Count);
        int j = random.Next(0, route.Count);

        while (i == j)
        {
            j = random.Next(0, route.Count);
        }

        if (i > j)
        {
            int temp = i;
            i = j;
            j = temp;
        }

        List<int> neighbor = new List<int>(route);
        double cost = double.MaxValue;

        if (neighborType == "reverse")
        {
            neighbor.Reverse(i, j - i + 1);
            cost = CalculateRouteCost(neighbor, distances);
        }
        else if (neighborType == "swap")
        {
            (neighbor[i], neighbor[j]) = (neighbor[j], neighbor[i]);
            cost = CalculateRouteCost(neighbor, distances);
        }
        else if (neighborType == "insert")
        {
            int city = neighbor[i];
            neighbor.RemoveAt(i);
            neighbor.Insert(j, city);
            cost = CalculateRouteCost(neighbor, distances);
        }

        return (neighbor, cost);
    }
}
