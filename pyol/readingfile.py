functions = []

for k in range(3):
    def f(k=k):
        return k

    # alternatively: f = lambda: i

    functions.append(f)
print(f)