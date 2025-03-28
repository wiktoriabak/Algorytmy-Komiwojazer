import random
import math
import time
import pandas as pd

def load_distance_matrix(file_path):
    data = pd.ExcelFile(file_path)
    sheet_data = data.parse(data.sheet_names[0])
    distance_matrix = sheet_data.iloc[:, 1:].values
    return distance_matrix

def route_length(route, distance_matrix):
    length = 0
    for i in range(len(route) - 1):
        length += distance_matrix[route[i]][route[i + 1]]  # odległość między kolejnymi miastami
    length += distance_matrix[route[-1]][route[0]]  # odległość z ostatniego miasta do pierwszego
    return length

def generate_neighbour(route, neighbourhood_type):
    if neighbourhood_type == 'swap':
        return swap_move(route)
    elif neighbourhood_type == 'insert':
        return insert_move(route)
    elif neighbourhood_type == 'reverse':
        return reverse_move(route)

# Funkcja swap (zamiana dwóch miast)
def swap_move(route):
    new_route = route[:]  # Tworzymy kopię trasy
    i, j = random.sample(range(len(route)), 2)  # Wybieramy dwa różne indeksy
    new_route[i], new_route[j] = new_route[j], new_route[i]  # Zamiana miejscami
    return new_route

# Funkcja insert (przenoszenie miasta na inne miejsce)
def insert_move(route):
    new_route = route[:]  # Tworzymy kopię trasy
    i, j = random.sample(range(len(route)), 2)  # Wybieramy dwa różne indeksy
    city = new_route.pop(i)  # Usuwamy miasto z indeksu i
    new_route.insert(j, city)  # Wstawiamy miasto na nowym miejscu
    return new_route

# Funkcja reverse (odwracanie fragmentu trasy)
def reverse_move(route):
    new_route = route[:]  # Tworzymy kopię trasy
    i, j = random.sample(range(len(route)), 2)  # Wybieramy dwa różne indeksy
    if i > j:  # Upewniamy się, że i < j, aby nie wystąpiły błędy
        i, j = j, i
    new_route[i:j + 1] = reversed(new_route[i:j + 1])  # Odwracamy fragment trasy
    return new_route

# Algorytm symulowanego wyżarzania
def simulated_annealing(distance_matrix, initial_temperature, alpha, n_iter_per_temp, max_iter_without_improvement,
                        neighbourhood_type):
    n_cities = len(distance_matrix)
    current_solution = list(range(n_cities))  # Tworzymy pełną listę miast
    random.shuffle(current_solution)  # Tasowanie
    current_length = route_length(current_solution, distance_matrix)

    best_solution = current_solution[:]
    best_length = current_length

    temperature = initial_temperature
    iter_without_improvement = 0
    iterations = 0
    all_lengths = []

    start_time = time.time()

    while temperature > 1:
        for _ in range(n_iter_per_temp):
            new_solution = generate_neighbour(current_solution, neighbourhood_type)
            new_length = route_length(new_solution, distance_matrix)
            delta = new_length - current_length

            if delta < 0 or random.random() < math.exp(-delta / temperature):
                current_solution = new_solution
                current_length = new_length
                iter_without_improvement = 0
                if current_length < best_length:
                    best_solution = current_solution[:]
                    best_length = current_length
            else:
                iter_without_improvement += 1

            if iter_without_improvement >= max_iter_without_improvement:
                print(f"Zatrzymano po {iterations} iteracjach z brakiem poprawy.")
                break

            all_lengths.append(current_length)

        temperature *= alpha
        iterations += 1

        if iter_without_improvement >= max_iter_without_improvement:
            break

    elapsed_time = time.time() - start_time

    avg_length = sum(all_lengths) / len(all_lengths) if all_lengths else 0

    return best_solution, best_length, avg_length, elapsed_time

# Funkcja zapisująca wyniki do pliku Excel
def save_results_to_excel(results, filename='TSP_results_127.xlsx'):
    df = pd.DataFrame(results)
    df.to_excel(filename, index=False)

# Parametry eksperymentu
initial_temperatures = [500, 1000, 1500, 2000]  # 4 różne wartości temperatury początkowej
alphas = [0.85, 0.90, 0.95, 0.98]  # 4 różne wartości współczynnika redukcji temperatury
n_iter_per_temps = [50, 100, 150, 200]  # 4 różne wartości liczby iteracji dla jednej temperatury
max_iter_without_improvements = [20, 50, 100, 150]  # 4 różne wartości liczby iteracji bez poprawy
neighbourhood_types = ['swap', 'insert', 'reverse']  # 3 różne typy sąsiedztwa

file_path = "Dane_TSP_127.xlsx"  # Ścieżka do pliku z danymi
distance_matrix = load_distance_matrix(file_path)  # Wczytanie macierzy odległości

repetitions = 10  # Liczba powtórzeń eksperymentu

results = []

# Pętla wykonująca eksperymenty z różnymi parametrami
for T0 in initial_temperatures:
    for alpha in alphas:
        for n_iter_per_temp in n_iter_per_temps:
            for max_iter_without_improvement in max_iter_without_improvements:
                for neighbourhood_type in neighbourhood_types:
                    for rep in range(repetitions):  # Dodana pętla powtórzeń
                        best_solution, best_length, avg_length, elapsed_time = simulated_annealing(
                            distance_matrix,
                            T0,
                            alpha,
                            n_iter_per_temp,
                            max_iter_without_improvement,
                            neighbourhood_type
                        )

                        # Przesuń indeksy miast o 1 (do numeracji od 1 do 48), ale miasto 48 może być w dowolnym miejscu
                        best_solution = [city + 1 for city in best_solution]

                        results.append({
                            'Powtórzenie': rep + 1,  # Numer powtórzenia
                            'Temperatura Początkowa': T0,
                            'Redukcja Temperatury': alpha,
                            'Iteracje na Temperaturze': n_iter_per_temp,
                            'Maksymalna Liczba Iteracji bez Poprawy': max_iter_without_improvement,
                            'Rodzaj Sąsiedztwa': neighbourhood_type,
                            'Najlepsza Trasa': best_solution,
                            'Długość Najlepszej Trasy': best_length,
                            'Średnia Długość Trasy': avg_length,
                            'Czas Wykonania (s)': elapsed_time
                        })

# Zapisanie wyników do pliku Excel
save_results_to_excel(results)

print("Wyniki zapisane do pliku Excel.")
