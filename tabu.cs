using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OfficeOpenXml;

class TabuSearchTSP
{
    static void Main(string[] args)
    {
        string filePath = "C:\\Users\\iza51\\OneDrive\\Pulpit\\Studia\\Semestr 5\\Inteligencja obliczeniowa\\Dane_TSP_127.xlsx";
        double[,] distanceMatrix = null;

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //Licencja na pakiet
            distanceMatrix = LoadDistanceMatrix(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Błąd podczas wczytywania pliku: " + ex.Message);
            return;
        }

        var iterationsList = new List<int> { 100, 250, 500, 2000 };
        var noImprovementList = new List<int> { 25, 50, 100, 200 };
        var tabuListLengths = new List<int> { 5, 20, 50, 100 };
        var neighborhoods = new List<string> { "swap", "reverse", "insert" };

        List<Dictionary<string, object>> results = new List<Dictionary<string, object>>();

        foreach (var iterations in iterationsList)
        {
            foreach (var noImprovementLimit in noImprovementList)
            {
                foreach (var tabuListSize in tabuListLengths)
                {
                    foreach (var neighborhoodType in neighborhoods)
                    {
                        Stopwatch stopwatch = Stopwatch.StartNew();
                        List<int> solution = TabuSearch(distanceMatrix, tabuListSize, iterations, noImprovementLimit, neighborhoodType, 0.5);
                        double cost = CalculateRouteCost(solution, distanceMatrix);
                        stopwatch.Stop();

                        results.Add(new Dictionary<string, object>
                        {
                            { "iterations", iterations },
                            { "noImprovement", noImprovementLimit },
                            { "tabuLength", tabuListSize },
                            { "neighborhood", neighborhoodType },
                            { "bestPathLength", cost },
                            { "bestPath", string.Join(" ", solution.Select(city => city)) },
                            { "elapsedTime", stopwatch.Elapsed.TotalSeconds }
                        });
                    }
                }
            }
        }

        string newFilePath = "C:\\Users\\iza51\\OneDrive\\Pulpit\\Studia\\Semestr 5\\Inteligencja obliczeniowa\\Wyniki_127_10.xlsx";
        SaveResultsToExcel(results, newFilePath);

        Console.WriteLine("Wyniki zapisane do pliku Wyniki_kombinacje.xlsx.");
    }

    static double[,] LoadDistanceMatrix(string filePath)
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rows = worksheet.Dimension.Rows - 1;
            int cols = worksheet.Dimension.Columns - 1;
            double[,] matrix = new double[rows + 1, cols + 1]; 

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    matrix[i + 1, j + 1] = Convert.ToDouble(worksheet.Cells[i + 2, j + 2].Value); // Zaczynamy od 1
                }
            }

            return matrix;
        }
    }

    static void SaveResultsToExcel(List<Dictionary<string, object>> results, string filePath)
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets.Add("Results");
            worksheet.Cells[1, 1].Value = "Iterations";
            worksheet.Cells[1, 2].Value = "No Improvement";
            worksheet.Cells[1, 3].Value = "Tabu Length";
            worksheet.Cells[1, 4].Value = "Neighborhood";
            worksheet.Cells[1, 5].Value = "Best Path Length";
            worksheet.Cells[1, 6].Value = "Best Path";
            worksheet.Cells[1, 7].Value = "Elapsed Time (s)";

            int row = 2;
            foreach (var result in results)
            {
                worksheet.Cells[row, 1].Value = result["iterations"];
                worksheet.Cells[row, 2].Value = result["noImprovement"];
                worksheet.Cells[row, 3].Value = result["tabuLength"];
                worksheet.Cells[row, 4].Value = result["neighborhood"];
                worksheet.Cells[row, 5].Value = result["bestPathLength"];
                worksheet.Cells[row, 6].Value = result["bestPath"];
                worksheet.Cells[row, 7].Value = result["elapsedTime"];
                row++;
            }

            package.Save();
        }
    }

    static List<int> TabuSearch(double[,] distanceMatrix, int tabuListSize, int maxIterations, int noImprovementLimit, string neighborhoodType, double randomnessFactor)
    {
        int n = distanceMatrix.GetLength(0) - 1;
        Random rand = new Random();
        List<int> currentSolution = GenerateHybridSolution(distanceMatrix, randomnessFactor, rand);
        List<int> bestSolution = new List<int>(currentSolution);
        double bestCost = CalculateRouteCost(bestSolution, distanceMatrix);// najlepsze globalne rozwiazanie 

        Queue<List<int>> tabuList = new Queue<List<int>>();
        int noImprovementCounter = 0;

        for (int iteration = 0; iteration < maxIterations; iteration++)
        {
            List<List<int>> neighborhood = GenerateNeighborhood(currentSolution, neighborhoodType);
            List<int> bestNeighbor = null;
            double bestNeighborCost = double.MaxValue;

            foreach (var neighbor in neighborhood)
            {
                if (!IsTabu(tabuList, neighbor)) //Czy nie jest na liscie tabu, jeśli nie jest to warunek
                {
                    double cost = CalculateRouteCost(neighbor, distanceMatrix);
                    if (cost < bestNeighborCost) //pierwszy sasiad zawsze spelni warunek
                    {
                        bestNeighbor = neighbor;
                        bestNeighborCost = cost;
                    }
                }
            }

            if (bestNeighbor != null) //jesli znalezniono lepszego sasiada 
            {
                currentSolution = bestNeighbor;
                if (tabuList.Count >= tabuListSize) tabuList.Dequeue(); //aktualizacja listy tabu, jesli jest pelna to element usuwany
                tabuList.Enqueue(new List<int>(currentSolution));

                if (bestNeighborCost < bestCost) //best solution to droga pierwszego currentr solution
                {
                    bestSolution = new List<int>(bestNeighbor); //dodanie nowego najlepszego rozwiazania 
                    bestCost = bestNeighborCost;
                    noImprovementCounter = 0;
                }
                else
                {
                    noImprovementCounter++;
                }
            }

            if (noImprovementCounter >= noImprovementLimit)
                break;
        }

        return bestSolution;
    }

    static double CalculateRouteCost(List<int> solution, double[,] distanceMatrix)
    {
        double cost = 0;
        for (int i = 0; i < solution.Count - 1; i++)
        {
            cost += distanceMatrix[solution[i], solution[i + 1]];
        }
        cost += distanceMatrix[solution[^1], solution[0]]; // Dodaj koszt powrotu do początkowego miasta
        return cost;
    }

    static List<int> GenerateHybridSolution(double[,] distanceMatrix, double randomnessFactor, Random rand)
    {
        int n = distanceMatrix.GetLength(0) - 1;
        List<int> solution = new List<int>();
        HashSet<int> visited = new HashSet<int>();

        int currentCity = rand.Next(1, n + 1);
        solution.Add(currentCity);
        visited.Add(currentCity);

        while (solution.Count < n)
        {
            if (rand.NextDouble() < randomnessFactor)
            {
                int nextCity;
                do
                {
                    nextCity = rand.Next(1, n + 1);
                } while (visited.Contains(nextCity));
                solution.Add(nextCity);
                visited.Add(nextCity);
            }
            else
            {
                double minDistance = double.MaxValue;
                int nextCity = -1;
                foreach (var city in Enumerable.Range(1, n).Where(c => !visited.Contains(c)))
                {
                    if (distanceMatrix[currentCity, city] < minDistance)
                    {
                        minDistance = distanceMatrix[currentCity, city];
                        nextCity = city;
                    }
                }
                solution.Add(nextCity);
                visited.Add(nextCity);
            }
        }

        return solution;
    }

    static List<List<int>> GenerateNeighborhood(List<int> solution, string neighborhoodType)
    {
        List<List<int>> neighborhood = new List<List<int>>();
        for (int i = 0; i < solution.Count; i++)
        {
            for (int j = i + 1; j < solution.Count; j++)
            {
                List<int> neighbor = new List<int>(solution);
                switch (neighborhoodType)
                {
                    case "swap":
                        (neighbor[i], neighbor[j]) = (neighbor[j], neighbor[i]);
                        break;
                    case "reverse":
                        neighbor.Reverse(i, j - i + 1);
                        break;
                    case "insert":
                        int city = neighbor[j];
                        neighbor.RemoveAt(j);
                        neighbor.Insert(i, city);
                        break;
                }
                neighborhood.Add(neighbor);
            }
        }
        return neighborhood;
    }

    static bool IsTabu(Queue<List<int>> tabuList, List<int> solution)
    {
        foreach (var tabuSolution in tabuList)
        {
            if (Enumerable.SequenceEqual(tabuSolution, solution))
                return true;
        }
        return false;
    }
}
