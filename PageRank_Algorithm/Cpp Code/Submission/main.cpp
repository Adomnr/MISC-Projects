#include <iostream>
#include <fstream>
#include <sstream>
#include <vector>
#include <unordered_map>
#include <cmath>
#include <chrono>  

using namespace std;
using namespace std::chrono;  

const double DAMPING_FACTOR = 0.85;
const int MAX_ITERATIONS = 100;
const double TOLERANCE = 1e-6;

double max_rank = 0;

unordered_map<int, vector<int>> loadGraph(const string& filename, int& numNodes) {
    ifstream file(filename);
    string line;
    unordered_map<int, vector<int>> adjList;
    numNodes = 0;

    while (getline(file, line)) {
        stringstream ss(line);
        string node1, node2, weight;
        getline(ss, node1, ',');
        getline(ss, node2, ',');
        getline(ss, weight, ',');  // Ignoring the weight, assuming unweighted graph

        try {
            int from = stoi(node1);
            int to = stoi(node2);

            adjList[from].push_back(to);

            numNodes = max(numNodes, max(from, to));
        } catch (const invalid_argument& e) {
            cerr << "Error: Invalid data encountered in line: " << line << endl;
            continue;
        } catch (const out_of_range& e) {
            cerr << "Error: Number out of range in line: " << line << endl;
            continue;
        }
    }
    numNodes++; // Assuming nodes are 0-indexed
    return adjList;
}

vector<double> initializePageRank(int numNodes) {
    vector<double> pagerank(numNodes, 1.0 / numNodes);
    return pagerank;
}

vector<double> calculatePageRank(const unordered_map<int, vector<int>>& adjList, int numNodes) {
    vector<double> pagerank = initializePageRank(numNodes);
    vector<double> newPagerank(numNodes, 0.0);

    for (int iter = 0; iter < MAX_ITERATIONS; iter++) {
        fill(newPagerank.begin(), newPagerank.end(), (1.0 - DAMPING_FACTOR) / numNodes);

        for (const auto& [from, toNodes] : adjList) {
            double outboundRank = pagerank[from] / toNodes.size();
            for (int to : toNodes) {
                newPagerank[to] += DAMPING_FACTOR * outboundRank;
            }
        }

        // Convergence check
        double diff = 0.0;
        for (int i = 0; i < numNodes; i++) {
            diff += fabs(newPagerank[i] - pagerank[i]);
        }

        pagerank = newPagerank;

        if (diff < TOLERANCE) {
            break;
        }
    }

    return pagerank;
}


int main() {
    // Start timing
    auto start = high_resolution_clock::now();

    int numNodes;
    string filename = "pagerank_graph_1000_websites_edges.csv";
    unordered_map<int, vector<int>> adjList = loadGraph(filename, numNodes);

    vector<double> pagerank = calculatePageRank(adjList, numNodes);

    for (int i = 0; i < numNodes; i++) {
        cout << "Node " << i << ": " << pagerank[i] << endl;
    }

    auto end = high_resolution_clock::now();
    auto duration = duration_cast<milliseconds>(end - start);

    // Output the time taken
    cout << "Execution time: " << duration.count() << " milliseconds" << endl;

    return 0;
}

