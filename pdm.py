ln = '10;12-19;358-527;529-547;549-621;623-655;657;659;661-671;673;675-677;680;682'
ln_1 = ln.split(';')
res2 = []
for i in ln_1:
    if '-' in i:
	    exp = i.split('-')
	    l1 = (range(int(exp[0]), int(exp[1])+1))
	    for n in l1:
		    res2.append(n)
    else:
	    res2.append(int(i))
pdm = [436, 438, 439, 442, 443, 444, 445, 446, 447, 449, 450, 451, 452, 453, 454, 455, 456, 458, 459, 461, 462, 463, 464, 465, 466, 467, 468, 469, 471, 472, 473, 475, 477, 478, 479, 480, 481, 483, 484, 485, 486, 487, 488, 490, 491, 492, 493, 494, 496, 497, 498, 500, 501, 502, 503, 504, 507, 509, 510, 512, 514, 515, 516, 518, 566, 567, 569, 570, 571, 572, 574, 575, 576, 577, 578, 580, 581, 582, 584, 585, 586, 587, 588, 589, 590, 591, 593, 594, 596, 597, 598, 600, 601, 602, 603, 604, 605, 606, 607, 608, 610, 611, 612, 613, 614, 615, 616, 617, 619, 620, 621, 623, 624, 626, 628, 629, 630, 631, 632, 633, 634, 635, 636, 637, 638, 639, 640, 641, 642, 643, 644, 646, 647, 648, 649, 650, 651, 653, 654, 655, 657, 659, 661, 662, 663, 664, 666, 667, 668, 669, 670, 671, 673, 675, 676, 677, 680, 682]
missed = []
for j in res2:
    if j not in pdm:
	    missed.append(j)
print missed

	    