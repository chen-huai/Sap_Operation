global a
a = 1


def op1(func):
    def warraper(*args, **kwargs):
        a = 0
        func(*args, **kwargs)

    return warraper


@op1
def op3(a):
    a += 1
    print(3, a)


op3(a)
