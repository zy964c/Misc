x = {'c': [3, 2], 'a': [4, 0], 'b': [0, 43]}
sorted_x = sorted(x.items(), key=lambda mysum: mysum[0])
print sorted_x
sorted_x1 = dict(sorted_x)
sorted_x2 = sorted(sorted_x1.items(), key=lambda mysum: sum(mysum[1]))
print sorted_x2
