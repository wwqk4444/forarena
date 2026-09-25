#!/usr/bin/env python3
"""Random test generator for KRIPTOGRAM.

The sentence is generated first, then the text is built so that the sentence is
guaranteed to occur (otherwise the "solution always exists" promise breaks).
"""
import random
import sys

seed = int(sys.argv[1])
random.seed(seed)

ALPHA = random.randint(1, 4)          # tiny alphabet -> many repeated words
words = [chr(ord('a') + i) for i in range(ALPHA)]

m = random.randint(1, 12)
sentence = [random.choice(words) for _ in range(m)]

n = random.randint(m, 30)
text = [random.choice(words) for _ in range(n)]
pos = random.randint(0, n - m)        # plant the sentence at a random position
text[pos:pos + m] = sentence[:]

# random renamings on both sides so that the words themselves differ
def rename(seq):
    perm = list(words)
    random.shuffle(perm)
    table = {w: perm[i] for i, w in enumerate(words)}
    return [table[w] for w in seq]

text = rename(text)
sentence = rename(sentence)

print(' '.join(text) + ' $')
print(' '.join(sentence) + ' $')
