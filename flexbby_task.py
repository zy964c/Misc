def foo(a={}):
    a[len(a)] = 0
    print a

foo()
foo()
foo({})

#>>> foo()
#{0: 0}
#>>> foo()
#{0: 0, 1: 0}
#>>> foo({})
#{0: 0}
#>>> 