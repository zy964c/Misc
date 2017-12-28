path_released = 'C:\Temp\zy964c\lighting_all.txt'
path_seed = 'C:\Temp\zy964c\seed_all.txt'

def parse_data(file_path, return_type):
    seed = open(file_path, 'r')
    seed_data  =  list(seed)
    seed_filtered = []
    for n in seed_data:
        upd1 = n.replace(' ', '').replace('\n', '').replace('"', '')
        upd2 = upd1.split(",")
        upd = upd2[:3]
        seed_filtered.append(upd)

    part_numbers = []
    revisions = []
    names = []
    for elem in seed_filtered:
        pn = elem[1]
        part_numbers.append(pn)
        rev = elem[0]
        revisions.append(rev)
        name = elem[2]
        names.append(name)

    if return_type == 'rev':
        dict1 = dict(zip(part_numbers, revisions))
    else:
        dict1 = dict(zip(part_numbers, names))

    return dict1


def check(dict_released, dict_seed, dict_released_names):

    to_create = 0
    to_upd_discrptn = 0
    dict_released_sorted_keys = dict_released_names.keys()
    dict_released_sorted_keys.sort()
    for key in dict_released_sorted_keys:
        if dict_seed.has_key(key):
            if dict_released[key] == dict_seed[key]:
                continue
            else:
                print key + ' has a SEED model but Rev.' + dict_seed[key] + ' should be updated to Rev.' + dict_released[key]
                to_upd_discrptn += 1
        else:
            #print key + ' Rev.' + dict_released[key] + ' (Description: ' + dict_released_names[key] + ')' + " - NO SEED MODEL FOUND"
            print key
    for key in dict_released_sorted_keys:
        if dict_seed.has_key(key):
            if dict_released[key] == dict_seed[key]:
                continue
            else:
                print key + ' has a SEED model but Rev.' + dict_seed[key] + ' should be updated to Rev.' + dict_released[key]
                to_upd_discrptn += 1
        else:
            #print key + ' Rev.' + dict_released[key] + ' (Description: ' + dict_released_names[key] + ')' + " - NO SEED MODEL FOUND"
            print dict_released[key]
    for key in dict_released_sorted_keys:
        if dict_seed.has_key(key):
            if dict_released[key] == dict_seed[key]:
                continue
            else:
                print key + ' has a SEED model but Rev.' + dict_seed[key] + ' should be updated to Rev.' + dict_released[key]
                to_upd_discrptn += 1
        else:
            #print key + ' Rev.' + dict_released[key] + ' (Description: ' + dict_released_names[key] + ')' + " - NO SEED MODEL FOUND"
            print dict_released_names[key]
            to_create += 1
    if to_upd_discrptn > 0:
        print 'SEED models to update version: ' + str(to_upd_discrptn)
    if to_create > 0:
        print 'SEED models to create: ' + str(to_create)
            

dict_released = parse_data(path_released, 'rev')
dict_seed = parse_data(path_seed, 'rev')
dict_released_names = parse_data(path_released, 'name')
dict_seed_names = parse_data(path_seed, 'name')


check(dict_released, dict_seed, dict_released_names)

dict_seed_sorted_keys = dict_seed_names.keys()
dict_seed_sorted_keys.sort()
a = 0
for key in dict_seed_sorted_keys:
    print key
    a += 1
for key in dict_seed_sorted_keys:
    print dict_seed_names[key]
for key in dict_seed_sorted_keys:
    if dict_seed[key] != '---':
        to_print = dict_seed[key]
        print to_print[-1]
    else:
        print dict_seed[key]
print a

        
