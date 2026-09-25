// COCI 2011/2012, 4. kolo, task KRIPTOGRAM
// Find the first occurrence, inside the encrypted message, of a window of
// consecutive words whose "equality pattern" is the same as the known sentence.
#include <bits/stdc++.h>
using namespace std;

static vector<string> readWords(istream &in) {
    string line;
    if (!getline(in, line)) return {};
    vector<string> w;
    string cur;
    for (char c : line) {
        if (c == '$') break;                       // '$' ends the message
        if (c == ' ' || c == '\t' || c == '\r') {
            if (!cur.empty()) { w.push_back(cur); cur.clear(); }
        } else {
            cur.push_back(c);
        }
    }
    if (!cur.empty()) w.push_back(cur);
    return w;
}

// d[i] = i - (previous occurrence of the same word), or INF if there is none.
static vector<int> encode(const vector<string> &w) {
    const int INF = INT_MAX / 2;
    vector<int> d(w.size(), INF);
    unordered_map<string, int> last;
    last.reserve(w.size() * 2 + 1);
    for (int i = 0; i < (int)w.size(); ++i) {
        auto it = last.find(w[i]);
        if (it != last.end()) d[i] = i - it->second;
        last[w[i]] = i;
    }
    return d;
}

int main() {
    ios_base::sync_with_stdio(false);
    cin.tie(nullptr);

    vector<string> text = readWords(cin);
    vector<string> pat = readWords(cin);
    int n = (int)text.size(), m = (int)pat.size();

    vector<int> a = encode(text), b = encode(pat);

    // Does text position i fit pattern position j (0-based), given that the
    // currently matched prefix of the pattern has length j+1?
    auto eq = [&](int j, const vector<int> &txt, int i) {
        int cap = j + 1;                 // length of the window that is compared
        return min(txt[i], cap) == min(b[j], cap);
    };

    vector<int> fail(m + 1, 0);         // KMP prefix function of the pattern
    for (int i = 1, k = 0; i < m; ++i) {
        while (k > 0 && !eq(k, b, i)) k = fail[k];
        if (eq(k, b, i)) ++k;
        fail[i + 1] = k;
    }

    int answer = 0;
    for (int i = 0, k = 0; i < n; ++i) {
        while (k > 0 && !eq(k, a, i)) k = fail[k];
        if (eq(k, a, i)) ++k;
        if (k == m) { answer = i - m + 2; break; }   // 1-based position
    }
    cout << answer << '\n';
    return 0;
}
