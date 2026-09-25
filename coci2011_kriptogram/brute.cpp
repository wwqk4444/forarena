// Brute force reference: try every window, compare the equality patterns.
#include <bits/stdc++.h>
using namespace std;

static vector<string> readWords(istream &in) {
    string line;
    if (!getline(in, line)) return {};
    vector<string> w;
    string cur;
    for (char c : line) {
        if (c == '$') break;
        if (c == ' ' || c == '\t' || c == '\r') {
            if (!cur.empty()) { w.push_back(cur); cur.clear(); }
        } else cur.push_back(c);
    }
    if (!cur.empty()) w.push_back(cur);
    return w;
}

int main() {
    auto T = readWords(cin), P = readWords(cin);
    int n = T.size(), m = P.size();
    for (int s = 0; s + m <= n; ++s) {
        bool ok = true;
        for (int k = 0; k < m && ok; ++k)
            for (int l = 0; l < k && ok; ++l)
                if ((T[s + k] == T[s + l]) != (P[k] == P[l])) ok = false;
        if (ok) { cout << s + 1 << '\n'; return 0; }
    }
    cout << 0 << '\n';
    return 0;
}
