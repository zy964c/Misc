class node(object):
    def __init__(self, data):
        self.data = data
        self.next = None
        self.prev = None

class linked_list(object):
    def __init__(self):
        self.first = None
        self.last = None

    def add(self, node):
        if self.first == None:
            self.first = node
            self.last = node
        else:
            self.first.prev = node
            node.next = self.first
            self.first = node

    def reverse(self):
        if self.first == None:
            return
        self.first, self.last = self.last, self.first
        cur_node = self.first
        while True:
            cur_node.next, cur_node.prev = cur_node.prev, cur_node.next
            if cur_node.next is None:
                break
            cur_node = cur_node.next
            

if __name__ == "__main__":

    a = linked_list()
    n1 = node('a')
    n2 = node('b')
    n3 = node('c')
    a.add(n1)
    a.add(n2)
    a.add(n3)
    print a.first.data
    print a.first.next.data
    print a.first.next.next.data
    a.reverse()
    print a.first.data
    print a.first.next.data
    print a.first.next.next.data
    a.reverse()
    print a.first.data
    print a.first.next.data
    print a.first.next.next.data
    
