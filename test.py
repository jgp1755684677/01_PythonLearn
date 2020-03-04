def test(a):
    a = a + 1
    return a


if __name__ == '__main__':
    a = 1
    a = test(a)
    print(a)