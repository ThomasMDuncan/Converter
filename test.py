import threading, queue, time

# q = queue.Queue(maxsize=3)

# def worker():
#     while True:
#         item = q.get()
#         print(f'Working on {item}')
#         time.sleep(3)
#         print(f'Finished {item}')
#         q.task_done()

# def stuff():
#     pass

# for n in range(10):
#     q.put(n)

# # turn-on the worker thread
# threading.Thread(target=worker, daemon=True).start()

# # send thirty task requests to the worker
# for item in range(10):
#     q.put(item)
# print('All task requests sent\n', end='')

# # block until all tasks are done
# q.join()
# print('All work completed')

#New

# def do_work(item):
#     time.sleep(2)

# def worker():
#     while True:
#         item = q.get()
#         print(f'Working on {item}')
#         do_work(item)
#         print(f'Done working on {item}')
#         q.task_done()

# print("Start")

# q = queue.Queue()
# for i in range(4):
#      t = threading.Thread(target=worker)
#      t.daemon = True
#      t.start()

# for item in range(15):
#     q.put(item)

# q.join() 
# print('All Done')

class tes():
    def __init__(self) -> None:
        self.__stuff = 0
        self.other = "new"

    def meth(self):
        print(self.other)
        print(self.__stuff)
t = tes()
t.meth()