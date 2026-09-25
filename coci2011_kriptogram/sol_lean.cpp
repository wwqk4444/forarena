// COCI 2011/2012, 4. kolo, task KRIPTOGRAM  (memory-lean version)
//
// The known sentence and the encrypted message are equal "up to renaming of
// words", so a window of the message matches iff the two word sequences have
// the same equality pattern (same words equal, different words different).
#include <bits/stdc++.h>
using namespace std;

struct Line {
    string buf;                       // owns the characters
    vector<string_view> word;         // views into buf
};

static Line readLine(istream &in) {
    Line L;
    getline(in, L.buf);
    size_t i = 0, n = L.buf.size();
    while (i < n) {
        if (L.buf[i] == '$') break;                    // '$' ends the line
        if (L.buf[i] == ' ') { ++i; continue; }
        size_t j = i;
        while (j < n && L.buf[j] != ' ' && L.buf[j] != '$') ++j;
        L.word.emplace_back(L.buf.data() + i, j - i);
        i = j;
    }
    return L;
}

// d[i] = distance to the previous occurrence of the same word, or INF.
static vector<int> encode(const vector<string_view> &w) {
    const int INF = INT_MAX / 2;
    vector<int> d(w.size(), INF);
    unordered_map<string_view, int> last;
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

    Line text = readLine(cin), pat = readLine(cin);
    int n = text.word.size(), m = pat.word.size();
    if (m == 0) { cout << 0 << '\n'; return 0; }

    vector<int> a = encode(text.word), b = encode(pat.word);

    // Text position i against pattern position j while a prefix of length j
    // has already been matched: only repetitions inside the window count.
    auto eq = [&](int j, const vector<int> &txt, int i) {
        int cap = j + 1;
        return min(txt[i], cap) == min(b[j], cap);
    };

    vector<int> fail(m + 1, 0);
    for (int i = 1, k = 0; i < m; ++i) {
        while (k > 0 && !eq(k, b, i)) k = fail[k];
        if (eq(k, b, i)) ++k;
        fail[i + 1] = k;
    }

    int answer = 0;
    for (int i = 0, k = 0; i < n; ++i) {
        while (k > 0 && !eq(k, a, i)) k = fail[k];
        if (eq(k, a, i)) ++k;
        if (k == m) { answer = i - m + 2; break; }
    }
    cout << answer << '\n';
    return 0;
}
