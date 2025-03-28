using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

class NearestNeighbor
{
    static void Main(string[] args)
    {
        string inputFilePath = "C:\\Users\\laura_a1y0snp\\Desktop\\5 semestr\\IO\\Dane_TSP_48.xlsx";
        string outputFilePath = "C:\\Users\\laura_a1y0snp\\Desktop\\5 semestr\\IO\\NN_48.xlsx";

 
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        double[,] distanceMatrix = LoadDistanceMatrix(inputFilePath);

        int cityCount = distanceMatrix.GetLength(0);
        double bestDistance = double.MaxValue;
        List<int> bestPath = null;
        double totalDistance = 0;

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("NN_Results");

            worksheet.Cell(1, 1).Value = "Start City";
            worksheet.Cell(1, 2).Value = "Best Distance";

            int row = 2;

            for (int startCity = 0; startCity < cityCount; startCity++)
            {
                List<int> path = FindShortestPath(distanceMatrix, startCity);
                double pathDistance = CalculatePathDistance(distanceMatrix, path);

                totalDistance += pathDistance;

                worksheet.Cell(row, 1).Value = startCity + 1; 
                worksheet.Cell(row, 2).Value = pathDistance;
                row++;

                if (pathDistance < bestDistance)
                {
                    bestDistance = pathDistance;
                    bestPath = new List<int>(path);
                }
            }

            double averageDistance = totalDistance / cityCount;

            worksheet.Cell(row + 1, 1).Value = "Best Path";
            worksheet.Cell(row + 1, 2).Value = string.Join(" -> ", bestPath.Select(city => city + 1));
            worksheet.Cell(row + 2, 1).Value = "Shortest Distance";
            worksheet.Cell(row + 2, 2).Value = bestDistance;
            worksheet.Cell(row + 3, 1).Value = "Average Distance";
            worksheet.Cell(row + 3, 2).Value = averageDistance;

            stopwatch.Stop();
            TimeSpan elapsed = stopwatch.Elapsed;
            worksheet.Cell(row + 4, 1).Value = "Execution Time";
            worksheet.Cell(row + 4, 2).Value = elapsed.ToString();

            workbook.SaveAs(outputFilePath);
        }

        Console.WriteLine($"Wyniki zapisano do pliku: {outputFilePath}");
    }

    static double[,] LoadDistanceMatrix(string filePath)
    {
        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(1);

            int rowCount = worksheet.RowsUsed().Count();
            int colCount = worksheet.ColumnsUsed().Count();

            if (rowCount != colCount)
            {
                throw new InvalidOperationException("Macierz nie jest kwadratowa!");
            }

            double[,] matrix = new double[rowCount, colCount];

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    matrix[i, j] = worksheet.Cell(i + 1, j + 1).GetValue<double>();
                }
            }

            return matrix;
        }
    }

    static List<int> FindShortestPath(double[,] matrix, int startCity)
    {
        int cityCount = matrix.GetLength(0);
        bool[] visited = new bool[cityCount];
        List<int> path = new List<int>();

        int currentCity = startCity;
        path.Add(currentCity);
        visited[currentCity] = true;

        for (int step = 0; step < cityCount - 1; step++)
        {
            double minDistance = double.MaxValue;
            int nextCity = -1;

            for (int i = 0; i < cityCount; i++)
            {
                if (!visited[i] && matrix[currentCity, i] < minDistance)
                {
                    minDistance = matrix[currentCity, i];
                    nextCity = i;
                }
            }

            if (nextCity == -1)
            {
                throw new InvalidOperationException("Nie znaleziono drogi do kolejnego miasta.");
            }

            path.Add(nextCity);
            visited[nextCity] = true;
            currentCity = nextCity;
        }

        return path;
    }

    static double CalculatePathDistance(double[,] matrix, List<int> path)
    {
        double totalDistance = 0;
        for (int i = 0; i < path.Count - 1; i++)
        {
            totalDistance += matrix[path[i], path[i + 1]];
        }

        totalDistance += matrix[path[^1], path[0]];

        return totalDistance;
    }
}
