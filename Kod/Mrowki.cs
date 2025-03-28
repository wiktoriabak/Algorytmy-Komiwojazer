
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;

class AntColonyOptimization
{
    static void Main()
    {
        string excelFilePath = "C:\\Users\\laura_a1y0snp\\Desktop\\5 semestr\\IO\\Dane_TSP_127.xlsx";
        string outputExcelFilePath = "C:\\Users\\laura_a1y0snp\\Desktop\\ACO_Results.xlsx";

        double[,] distanceMatrix = ReadDistanceMatrixFromExcel(excelFilePath);

        const double alpha = 1.0;
        const double beta = 2.0; 
        const double evaporationRate = 0.5;

        List<int> numAntsList = new List<int> { 100, 200, 500, 700 };
        List<int> numIterationsList = new List<int> { 50, 100, 200, 500 };
        List<double> initialPheromoneList = new List<double> { 1.0, 2.0, 3.0, 4.0 };
        List<double> QList = new List<double> { 50.0, 100.0, 150.0, 200.0 };
        List<string> neighborhoodTypes = new List<string> { "reverse", "swap", "insert" };

        double globalBestDistance = double.MaxValue;
        int[] globalBestPath = null;
        string bestParameters = null;

        List<(string parameters, int[] path, double distance, double averageDistance)> localResults = new List<(string, int[], double, double)>();

        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        int totalCombinations = numAntsList.Count * numIterationsList.Count * initialPheromoneList.Count *
                        QList.Count * neighborhoodTypes.Count;
        int currentCombination = 0;

        int numOuterLoops = 20;

        foreach (string neighborhoodType in neighborhoodTypes)
        {
            foreach (int numAnts in numAntsList)
            {
                foreach (int numIterations in numIterationsList)
                {
                    foreach (double initialPheromone in initialPheromoneList)
                    {
                        foreach (double Q in QList)
                        {
                            currentCombination++;
                            int progress = (int)((currentCombination / (double)totalCombinations) * 100);
                            Console.Write($"\rPostęp: {progress}%");

                            double localBestDistance = double.MaxValue;
                            int[] localBestPath = null;

                            double totalDistanceForCurrentCombination = 0.0; 

                            for (int outerLoop = 0; outerLoop < numOuterLoops; outerLoop++)
                            {
                                double distance = RunACO(numAnts, numIterations, initialPheromone, alpha, beta, evaporationRate, Q, distanceMatrix, neighborhoodType, out int[] path);
                                if (distance < localBestDistance)
                                {
                                    localBestDistance = distance;
                                    localBestPath = path;
                                }
                                totalDistanceForCurrentCombination += distance; 
                            }

                            double averageDistanceForCurrentCombination = totalDistanceForCurrentCombination / numOuterLoops;

                            string parameters = $"neighborhoodType={neighborhoodType}, numAnts={numAnts}, numIterations={numIterations}, initialPheromone={initialPheromone}, Q={Q}";
                            localResults.Add((parameters, localBestPath, localBestDistance, averageDistanceForCurrentCombination));

                            Console.WriteLine($"\nParametry: {parameters}");
                            Console.WriteLine($"Najlepsza lokalna trasa: {string.Join(" -> ", localBestPath)}");
                            Console.WriteLine($"Najlepsza lokalna odległość: {localBestDistance}");
                            Console.WriteLine($"Średnia lokalna odległość: {averageDistanceForCurrentCombination}");

                            if (localBestDistance < globalBestDistance)
                            {
                                globalBestDistance = localBestDistance;
                                globalBestPath = localBestPath;
                                bestParameters = parameters;
                            }
                        }
                    }
                }
            }
        }

        stopwatch.Stop();
        TimeSpan elapsedTime = stopwatch.Elapsed;

        SaveResultsToExcel(localResults, outputExcelFilePath);

        Console.WriteLine("\n\nPodsumowanie globalne:");
        Console.WriteLine($"Najlepsza globalna ścieżka: {string.Join(" -> ", globalBestPath)}");
        Console.WriteLine($"Najlepsza globalna odległość: {globalBestDistance}");
        Console.WriteLine($"Najlepsze parametry: {bestParameters}");
        Console.WriteLine($"Całkowity czas wykonania: {elapsedTime}");
    }

    static void SaveResultsToExcel(List<(string parameters, int[] path, double distance, double averageDistance)> results, string filePath)
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet;

            var existingWorksheet = package.Workbook.Worksheets["Results"];
            if (existingWorksheet != null)
            {
                worksheet = existingWorksheet;
                worksheet.Cells.Clear();
            }
            else
            {
                worksheet = package.Workbook.Worksheets.Add("Results");
            }

            worksheet.Cells[1, 1].Value = "Parametry";
            worksheet.Cells[1, 2].Value = "Najlepsza lokalna trasa";
            worksheet.Cells[1, 3].Value = "Najlepsza lokalna odległość";
            worksheet.Cells[1, 4].Value = "Średnia lokalna odległość";

            int row = 2;
            foreach (var result in results)
            {
                worksheet.Cells[row, 1].Value = result.parameters;
                worksheet.Cells[row, 2].Value = string.Join(" -> ", result.path);
                worksheet.Cells[row, 3].Value = result.distance;
                worksheet.Cells[row, 4].Value = result.averageDistance;
                row++;
            }

            package.Save();
        }

        Console.WriteLine($"\nWyniki zapisane do pliku: {filePath}");
    }


    static double[,] ReadDistanceMatrixFromExcel(string filePath)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null || worksheet.Dimension == null)
            {
                throw new Exception("Nie znaleziono danych w pliku Excel.");
            }

            int numRows = worksheet.Dimension.Rows;
            int numCols = worksheet.Dimension.Columns;
            double[,] matrix = new double[numRows, numCols];

            for (int i = 1; i <= numRows; i++)
            {
                for (int j = 1; j <= numCols; j++)
                {
                    matrix[i - 1, j - 1] = worksheet.Cells[i, j].GetValue<double>();
                }
            }

            return matrix;
        }
    }

    static double RunACO(int numAnts, int numIterations, double initialPheromone, double alpha, double beta, double evaporationRate, double Q, double[,] distanceMatrix, string neighborhoodType, out int[] globalBestPath)
    {
        int numCities = distanceMatrix.GetLength(0);
        double globalBestDistance = double.MaxValue;
        globalBestPath = null;

        double[,] pheromones = new double[numCities, numCities];
        Random rand = new Random();

        for (int i = 0; i < numCities; i++)
        {
            for (int j = 0; j < numCities; j++)
            {
                pheromones[i, j] = initialPheromone;
            }
        }

        for (int iter = 0; iter < numIterations; iter++)
        {
            int[][] paths = new int[numAnts][];
            double[] distances = new double[numAnts];

            for (int k = 0; k < numAnts; k++)
            {
                int startCity = rand.Next(numCities);
                paths[k] = BuildPath(numCities, distanceMatrix, pheromones, alpha, beta, startCity, neighborhoodType);
                distances[k] = CalculateDistance(paths[k], distanceMatrix);

                if (distances[k] < globalBestDistance)
                {
                    globalBestDistance = distances[k];
                    globalBestPath = paths[k];
                }
            }

            UpdatePheromones(pheromones, paths, distances, Q, evaporationRate);
        }

        return globalBestDistance;
    }

    static int[] BuildPath(int numCities, double[,] distanceMatrix, double[,] pheromones, double alpha, double beta, int startCity, string neighborhoodType)
    {
        List<int> unvisited = new List<int>();
        for (int i = 0; i < numCities; i++) unvisited.Add(i);

        int[] path = new int[numCities];
        int currentCity = startCity;
        path[0] = currentCity;
        unvisited.Remove(currentCity);

        Random rand = new Random();

        for (int step = 1; step < numCities; step++)
        {
            double[] probabilities = new double[unvisited.Count];
            double sum = 0.0;

            for (int i = 0; i < unvisited.Count; i++)
            {
                int nextCity = unvisited[i];
                double pheromone = Math.Pow(pheromones[currentCity, nextCity], alpha);
                double heuristic = Math.Pow(1.0 / distanceMatrix[currentCity, nextCity], beta);
                probabilities[i] = pheromone * heuristic;
                sum += probabilities[i];
            }

            for (int i = 0; i < probabilities.Length; i++)
            {
                probabilities[i] /= sum;
            }

            double r = rand.NextDouble();
            double cumulative = 0.0;

            for (int i = 0; i < probabilities.Length; i++)
            {
                cumulative += probabilities[i];
                if (r <= cumulative)
                {
                    currentCity = unvisited[i];
                    break;
                }
            }

            path[step] = currentCity;
            unvisited.Remove(currentCity);
        }

        switch (neighborhoodType)
        {
            case "swap":
                ApplySwapNeighborhood(path, rand);
                break;
            case "reverse":
                ApplyReverseNeighborhood(path, rand);
                break;
            case "insert":
                ApplyInsertNeighborhood(path, rand);
                break;
        }

        return path;
    }

    static void ApplySwapNeighborhood(int[] path, Random rand)
    {
        int i = rand.Next(path.Length);
        int j = rand.Next(path.Length);
        (path[i], path[j]) = (path[j], path[i]);
    }

    static void ApplyReverseNeighborhood(int[] path, Random rand)
    {
        int i = rand.Next(path.Length);
        int j = rand.Next(i, path.Length);
        Array.Reverse(path, i, j - i + 1);
    }

    static void ApplyInsertNeighborhood(int[] path, Random rand)
    {
        int i = rand.Next(path.Length);
        int j = rand.Next(path.Length);
        int city = path[i];
        List<int> tempPath = new List<int>(path);
        tempPath.RemoveAt(i);
        tempPath.Insert(j, city);
        for (int k = 0; k < path.Length; k++) path[k] = tempPath[k];
    }

    static double CalculateDistance(int[] path, double[,] distanceMatrix)
    {
        double totalDistance = 0.0;
        for (int k = 0; k < path.Length - 1; k++)
        {
            totalDistance += distanceMatrix[path[k], path[k + 1]];
        }
        totalDistance += distanceMatrix[path[path.Length - 1], path[0]]; 
        return totalDistance;
    }

    static void UpdatePheromones(double[,] pheromones, int[][] paths, double[] distances, double Q, double evaporationRate)
    {
        int numCities = pheromones.GetLength(0);

        for (int i = 0; i < numCities; i++)
        {
            for (int j = 0; j < numCities; j++)
            {
                pheromones[i, j] *= (1.0 - evaporationRate);
            }
        }

        for (int k = 0; k < paths.Length; k++)
        {
            double contribution = Q / distances[k];
            for (int i = 0; i < paths[k].Length - 1; i++)
            {
                int from = paths[k][i];
                int to = paths[k][i + 1];
                pheromones[from, to] += contribution;
                pheromones[to, from] += contribution;
            }

            int last = paths[k][paths[k].Length - 1];
            int first = paths[k][0];
            pheromones[last, first] += contribution;
            pheromones[first, last] += contribution;
        }

    }
}