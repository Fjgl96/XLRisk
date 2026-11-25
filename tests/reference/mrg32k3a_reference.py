"""Reference values for the XLRisk self tests (Source/SelfTests.bas).

Re-implements, outside Excel, the exact arithmetic of Source/Rng.bas and the
Fisher-Yates shuffle of Source/Math.bas, prints every reference value used by
the VBA self tests and runs the same statistical checks.

Usage:  python tests/reference/mrg32k3a_reference.py
Needs only the Python standard library (scipy is used for the chi-square
critical values when it is installed).
"""
import math

M1 = 4294967087.0
M2 = 4294944443.0
A12, A13N, A21, A23N = 1403580.0, 810728.0, 527612.0, 1370589.0
NORM_DIVISOR = 4294967088.0
TWO_POW_32 = 4294967296.0


def dmod(a, b):
    r = a - math.floor(a / b) * b
    if r < 0:
        r += b
    if r >= b:
        r -= b
    return r


class Mrg32k3a:
    """Same Double arithmetic as Source/Rng.bas."""

    def __init__(self, state=None, seed=None):
        if seed is not None:
            self.seed(seed)
        else:
            self.s = [float(x) for x in state]

    def seed(self, seed):
        x = dmod(abs(math.floor(seed)), TWO_POW_32)
        s = []
        for i in range(6):
            x = dmod(69069.0 * x + 1.0, TWO_POW_32)
            s.append(dmod(x, M1 if i < 3 else M2))
        if s[0] == s[1] == s[2] == 0:
            s[0] = 12345.0
        if s[3] == s[4] == s[5] == 0:
            s[3] = 12345.0
        self.s = s

    def uniform(self):
        s = self.s
        p1 = A12 * s[1] - A13N * s[0]
        p1 -= math.floor(p1 / M1) * M1
        if p1 < 0:
            p1 += M1
        if p1 >= M1:
            p1 -= M1
        s[0], s[1], s[2] = s[1], s[2], p1
        p2 = A21 * s[5] - A23N * s[3]
        p2 -= math.floor(p2 / M2) * M2
        if p2 < 0:
            p2 += M2
        if p2 >= M2:
            p2 -= M2
        s[3], s[4], s[5] = s[4], s[5], p2
        return ((p1 - p2) if p1 > p2 else (p1 - p2 + M1)) / NORM_DIVISOR


def exact_reference(n):
    """L'Ecuyer's MRG32k3a with exact integer arithmetic, seeds 12345."""
    m1, m2 = 4294967087, 4294944443
    s = [12345] * 6
    out = []
    for _ in range(n):
        p1 = (1403580 * s[1] - 810728 * s[0]) % m1
        s[0], s[1], s[2] = s[1], s[2], p1
        p2 = (527612 * s[5] - 1370589 * s[3]) % m2
        s[3], s[4], s[5] = s[4], s[5], p2
        out.append(((p1 - p2) if p1 > p2 else (p1 - p2 + m1)) / (m1 + 1))
    return out


def chi2_critical(df, p=0.999):
    try:
        from scipy.stats import chi2
        return chi2.ppf(p, df)
    except ImportError:
        return None


def main():
    print("MRG32k3a, six seeds 12345 (published reference):")
    g = Mrg32k3a(state=[12345] * 6)
    for i, exact in enumerate(exact_reference(3), 1):
        print(f"  #{i}: {g.uniform()!r}   exact integer version: {exact!r}")

    print("RngSeed 12345:")
    g = Mrg32k3a(seed=12345)
    for i in range(1, 4):
        print(f"  #{i}: {g.uniform()!r}")

    # Uniformity, same checks as TestRngUniformity
    n, bins = 200000, 100
    g = Mrg32k3a(seed=2026)
    counts = [0] * bins
    su = su2 = slag = prev = 0.0
    for i in range(n):
        u = g.uniform()
        su += u
        su2 += u * u
        if i > 0:
            slag += u * prev
        counts[int(u * bins)] += 1
        prev = u
    mean = su / n
    var = su2 / n - mean * mean
    lag1 = (slag / (n - 1) - 0.25) * 12
    exp = n / bins
    chi = sum((c - exp) ** 2 / exp for c in counts)
    print(f"Uniformity (seed 2026): mean {mean:.6f} (tol {4*math.sqrt(1/12/n):.6f}), "
          f"variance {var:.6f} (exact {1/12:.6f}), lag-1 {lag1:.5f} (tol {4/math.sqrt(n):.5f}), "
          f"chi-square {chi:.2f} (critical {chi2_critical(bins - 1)})")

    # Shuffle, same checks as TestShuffle
    def shuffle(v, g, biased=False):
        j = len(v) - 1
        while j > 0:
            if biased:   # XLRisk 1.0: k = CLng(start + Rnd() * (J - start)), banker's rounding
                k = round(g.uniform() * j)
            else:        # XLRisk 1.1: k = start + Int(RngUniform() * (J - start + 1))
                k = int(g.uniform() * (j + 1))
            v[j], v[k] = v[k], v[j]
            j -= 1

    for biased in (False, True):
        g = Mrg32k3a(seed=99)
        pos = [0] * 10
        for _ in range(20000):
            v = list(range(1, 11))
            shuffle(v, g, biased)
            pos[v.index(1)] += 1
        exp = 20000 / 10
        chi = sum((c - exp) ** 2 / exp for c in pos)
        label = "XLRisk 1.0 shuffle" if biased else "XLRisk 1.1 shuffle"
        print(f"{label}: positions of value 1 {pos}, chi-square {chi:.2f} "
              f"(critical {chi2_critical(9)})")


if __name__ == "__main__":
    main()
