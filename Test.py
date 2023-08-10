class Test():
    def __init__(self):
        self.a = 1

    def op1(func,self):
        def warraper(*args, **kwargs):
            self.a = 0
            func(*args, **kwargs)
        return warraper

    def op2(func,self):
        def warraper(*args, **kwargs):
            self.a = 1
            func(*args, **kwargs)
        return warraper

    @op1
    def op3(self):
        self.a += 1
        print(3, self.a)

    @op2
    def op4(self):
        self.a += 1
        print(4, self.a)

    @op1
    def op5(self):
        self.a += 1
        print(5, self.a)

    @op2
    def op6(self):
        self.a += 1
        print(6, self.a)

    def op7(self):
        print(7, self.a)


test = Test()
test.op3()
test.op7()
test.op4()
test.op7()
test.op5()
test.op7()
test.op6()
test.op7()


import time


# def baiyu():
#     t1 = time.time()
#     print("我是攻城狮白玉")
#     time.sleep(2)
#     print("执行时间为：", time.time() - t1)
#
#
# def blog(name):
#     t1 = time.time()
#     print('进入blog函数')
#     name()
#     print('我的博客是 https://blog.csdn.net/zhh763984017')
#     print("执行时间为：", time.time() - t1)
#
#
# if __name__ == '__main__':
#     func = baiyu  # 这里是把baiyu这个函数名赋值给变量func
#     func()  # 执行func函数
#     print('------------')
#     blog(baiyu)  # 把baiyu这个函数作为参数传递给blog函数
#


# def count_time(func):
#     def wrapper():
#         t1 = time.time()
#         func()
#         print("执行时间为：", time.time() - t1)
#
#     return wrapper
#
# def baiyu():
#     print("我是攻城狮白玉")
#     time.sleep(2)
#
# if __name__ == '__main__':
#     baiyu = count_time(baiyu)  # 因为装饰器 count_time(baiyu) 返回的时函数对象 wrapper，这条语句相当于  baiyu = wrapper
#     baiyu()  # 执行baiyu()就相当于执行wrapper()


# import time
#
#
# def count_time(func):
#     def wrapper():
#         t1 = time.time()
#         func()
#         print("执行时间为：", time.time() - t1)
#
#     return wrapper
#
#
# @count_time
# def baiyu():
#     print("我是攻城狮白玉")
#     time.sleep(2)
#
#
# if __name__ == '__main__':
#     # baiyu = count_time(baiyu)  # 因为装饰器 count_time(baiyu) 返回的时函数对象 wrapper，这条语句相当于  baiyu = wrapper
#     # baiyu()  # 执行baiyu()就相当于执行wrapper()
#
#     baiyu()  # 用语法糖之后，就可以直接调用该函数了