using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace GeneticAlgorithmTSP
{
    class Program
    {
        static void Main(string[] args)
        {
            // Wczytaj macierz z pliku CSV
            string filePath = "C:\\Users\\malgo\\OneDrive\\Pulpit\\IO\\Dane_TSP_76.xlsx";
            int[,] distanceMatrix = LoadDistanceMatrix(filePath);

            // Parametry algorytmu
            int[] populationSizes = { 50, 100, 200, 500 };
            int[] maxIterationsArray = { 500, 1000, 2000, 5000 };

            string[] parentSelectionMethods = { "tournament", "roulette" };
            string[] mutationTypes = { "swap", "reverse", "insert" };
            string[] crossoverTypes = { "PMX", "OX" };

            // Wyniki zapisane do pliku
            string outputPath = "wyniki_testow76.csv";
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                writer.WriteLine("PopulationSize,MaxIterations,ParentSelection,MutationType,CrossoverType,BestFitness,Route");

                foreach (int populationSize in populationSizes)
                {
                    foreach (int maxIterations in maxIterationsArray)
                    {
                        foreach (string parentSelectionMethod in parentSelectionMethods)
                        {
                            foreach (string mutationType in mutationTypes)
                            {
                                foreach (string crossoverType in crossoverTypes)
                                {
                                    Console.WriteLine($"Rozpoczynam test: PopulationSize={populationSize}, MaxIterations={maxIterations}, ParentSelection={parentSelectionMethod}, MutationType={mutationType}, CrossoverType={crossoverType}");
                                    // Uruchom algorytm genetyczny dla danej kombinacji parametrów
                                    GeneticAlgorithm ga = new GeneticAlgorithm(
                                        distanceMatrix,
                                        populationSize,
                                        maxIterations,
                                        parentSelectionMethod,
                                        mutationType,
                                        crossoverType
                                    );

                                    var result = ga.Run();

                                    // Zapisz wyniki do pliku
                                    writer.WriteLine($"{populationSize},{maxIterations},{parentSelectionMethod},{mutationType},{crossoverType},{result.BestFitness},{string.Join(" ", result.BestRoute)}");
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine("Testy zakończone. Wyniki zapisane w pliku: " + outputPath);
        }

        static int[,] LoadDistanceMatrix(string filePath)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using var package = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0]; // Pobierz pierwszy arkusz

            int rows = worksheet.Dimension.Rows - 1; // Pomijamy pierwszy wiersz
            int cols = worksheet.Dimension.Columns - 1; // Pomijamy pierwszą kolumnę
            int[,] matrix = new int[rows, cols];

            for (int i = 2; i <= rows + 1; i++) // Wiersze od 2 (pomijając nagłówki)
            {
                for (int j = 2; j <= cols + 1; j++) // Kolumny od 2 (pomijając indeksy miast)
                {
                    string cellValue = worksheet.Cells[i, j].Text; // Pobierz wartość jako tekst
                    if (int.TryParse(cellValue, out int distance))
                    {
                        matrix[i - 2, j - 2] = distance; // Wstaw odległość do macierzy
                    }
                    else
                    {
                        throw new FormatException($"Nieprawidłowa wartość w komórce ({i},{j}): {cellValue}");
                    }
                }
            }

            return matrix;
        }

    }

    class GeneticAlgorithm
    {
        private readonly int[,] _distanceMatrix;
        private readonly int _populationSize;
        private readonly int _maxIterations;
        private readonly string _parentSelectionMethod;
        private readonly string _mutationType;
        private readonly string _crossoverType;

        private List<int[]> _population;
        private Random _random;

        public GeneticAlgorithm(int[,] distanceMatrix, int populationSize, int maxIterations, string parentSelectionMethod, string mutationType, string crossoverType)
        {
            _distanceMatrix = distanceMatrix;
            _populationSize = populationSize;
            _maxIterations = maxIterations;
            _parentSelectionMethod = parentSelectionMethod;
            _mutationType = mutationType;
            _crossoverType = crossoverType;
            _population = new List<int[]>();
            _random = new Random();
        }

        public (int BestFitness, int[] BestRoute) Run()
        {
            InitializePopulation();

            for (int iteration = 0; iteration < _maxIterations; iteration++)
            {
                List<int[]> newPopulation = new List<int[]>();

                while (newPopulation.Count < _populationSize)
                {
                    // Wybierz rodziców
                    var parents = SelectParents();

                    // Krzyżowanie
                    int[] offspring1, offspring2;
                    if (_random.NextDouble() < 0.8) // Domyślne prawdopodobieństwo krzyżowania
                    {
                        (offspring1, offspring2) = Crossover(parents.Item1, parents.Item2);
                    }
                    else
                    {
                        offspring1 = parents.Item1;
                        offspring2 = parents.Item2;
                    }

                    // Mutacja
                    Mutate(offspring1);
                    Mutate(offspring2);

                    newPopulation.Add(offspring1);
                    newPopulation.Add(offspring2);
                }

                // Aktualizacja populacji
                _population = newPopulation.OrderBy(CalculateFitness).Take(_populationSize).ToList();
            }

            // Zwrócenie najlepszego rozwiązania
            int[] bestSolution = _population.OrderBy(CalculateFitness).First();
            int bestFitness = CalculateFitness(bestSolution);
            return (bestFitness, bestSolution);
        }

        private void InitializePopulation()
        {
            int cities = _distanceMatrix.GetLength(0);
            for (int i = 0; i < _populationSize; i++)
            {
                int[] route = Enumerable.Range(0, cities).OrderBy(x => _random.Next()).ToArray();
                _population.Add(route);
            }
        }

        private (int[], int[]) SelectParents()
        {
            if (_parentSelectionMethod == "tournament")
            {
                return (TournamentSelection(), TournamentSelection());
            }
            else if (_parentSelectionMethod == "roulette")
            {
                return (RouletteSelection(), RouletteSelection());
            }

            throw new ArgumentException("Nieznana metoda selekcji rodziców: " + _parentSelectionMethod);
        }

        private int[] TournamentSelection()
        {
            int tournamentSize = 5;
            var candidates = _population.OrderBy(x => _random.Next()).Take(tournamentSize);
            return candidates.OrderBy(CalculateFitness).First();
        }

        private int[] RouletteSelection()
        {
            double totalFitness = _population.Sum(route => 1.0 / CalculateFitness(route));
            double randomValue = _random.NextDouble() * totalFitness;

            double cumulative = 0;
            foreach (var route in _population)
            {
                cumulative += 1.0 / CalculateFitness(route);
                if (cumulative >= randomValue)
                {
                    return route;
                }
            }

            return _population.Last();
        }

        private (int[], int[]) Crossover(int[] parent1, int[] parent2)
        {
            if (_crossoverType == "PMX")
            {
                return PMX(parent1, parent2);
            }
            else if (_crossoverType == "OX")
            {
                return OX(parent1, parent2);
            }

            throw new ArgumentException("Nieznana metoda krzyżowania: " + _crossoverType);
        }

        private void Mutate(int[] route)
        {
            if (_random.NextDouble() > 0.1) return; // Domyślne prawdopodobieństwo mutacji

            if (_mutationType == "swap")
            {
                int i = _random.Next(route.Length);
                int j = _random.Next(route.Length);
                (route[i], route[j]) = (route[j], route[i]);
            }
            else if (_mutationType == "reverse")
            {
                int i = _random.Next(route.Length);
                int j = _random.Next(route.Length);
                if (i > j) (i, j) = (j, i);
                Array.Reverse(route, i, j - i + 1);
            }
            else if (_mutationType == "insert")
            {
                int i = _random.Next(route.Length);
                int j = _random.Next(route.Length);
                int temp = route[i];
                if (i < j)
                {
                    Array.Copy(route, i + 1, route, i, j - i);
                }
                else
                {
                    Array.Copy(route, j, route, j + 1, i - j);
                }
                route[j] = temp;
            }
        }

        private int CalculateFitness(int[] route)
        {
            int fitness = 0;
            for (int i = 0; i < route.Length - 1; i++)
            {
                fitness += _distanceMatrix[route[i], route[i + 1]];
            }
            fitness += _distanceMatrix[route[^1], route[0]]; // Powrót do miasta początkowego
            return fitness;
        }

        private (int[], int[]) PMX(int[] parent1, int[] parent2)
        {
            int size = parent1.Length;
            int[] child1 = new int[size];
            int[] child2 = new int[size];
            Array.Fill(child1, -1);
            Array.Fill(child2, -1);

            int start = _random.Next(size);
            int end = _random.Next(size);
            if (start > end) (start, end) = (end, start);

            for (int i = start; i <= end; i++)
            {
                child1[i] = parent1[i];
                child2[i] = parent2[i];
            }

            for (int i = 0; i < size; i++)
            {
                if (child1[i] == -1)
                {
                    int gene = parent2[i];
                    while (child1.Contains(gene))
                    {
                        gene = parent2[Array.IndexOf(parent1, gene)];
                    }
                    child1[i] = gene;
                }

                if (child2[i] == -1)
                {
                    int gene = parent1[i];
                    while (child2.Contains(gene))
                    {
                        gene = parent1[Array.IndexOf(parent2, gene)];
                    }
                    child2[i] = gene;
                }
            }

            return (child1, child2);
        }

        private (int[], int[]) OX(int[] parent1, int[] parent2)
        {
            int size = parent1.Length;
            int[] child1 = new int[size];
            int[] child2 = new int[size];
            Array.Fill(child1, -1);
            Array.Fill(child2, -1);

            int start = _random.Next(size);
            int end = _random.Next(size);
            if (start > end) (start, end) = (end, start);

            Array.Copy(parent1, start, child1, start, end - start + 1);
            Array.Copy(parent2, start, child2, start, end - start + 1);

            int currentIndex1 = (end + 1) % size;
            int currentIndex2 = (end + 1) % size;

            for (int i = 0; i < size; i++)
            {
                int gene1 = parent2[(end + 1 + i) % size];
                int gene2 = parent1[(end + 1 + i) % size];

                if (!child1.Contains(gene1))
                {
                    child1[currentIndex1] = gene1;
                    currentIndex1 = (currentIndex1 + 1) % size;
                }

                if (!child2.Contains(gene2))
                {
                    child2[currentIndex2] = gene2;
                    currentIndex2 = (currentIndex2 + 1) % size;
                }
            }

            return (child1, child2);
        }
    }
}
