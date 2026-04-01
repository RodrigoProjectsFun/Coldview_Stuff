line1 = " 198771  00  I.N.T  604        23.00  SOL        23.00  SOL        23.00  AHO 194-36830982-0-10   14062025 234248 14062025 06-27"
line2 = "   3         6012 650036         PLIN-VANESSA BERROSPI     Visa Direct   PE    420829       NO   516523243643  05  NO 59E-COMMMERCE"

with open("ruler_output.txt", "w", encoding="utf-8") as f:
    for line in [line1, line2]:
        f.write(line + "\n")
        ruler10 = "".join([str(i // 10) if i % 10 == 0 else ' ' for i in range(len(line))])
        ruler = "".join([str(i % 10) for i in range(len(line))])
        f.write(ruler10 + "\n")
        f.write(ruler + "\n")
        f.write("-" * len(line) + "\n\n")
