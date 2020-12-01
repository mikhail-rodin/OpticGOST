import argparse
import os.path

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

plot_types={
    # name:(x_column, y_label,x_label,x_label_afocal, label)
    "lon": (1, r"$-m'$",             r"$\Delta s'$, мм", r"$\Delta s'$, дптр", u"Продольная сферическая аберрация"),
    "sph": (1, r"$m'-m_{гл}$, мм",   r"$\Delta y'$, мм",   r"$\sigma'$, мин", u"Поперечная сферическая аберрация"),
    "tang":(1, r"$m'-m_{гл}$, мм",   r"$\Delta y'$, мм",   r"$\sigma'$, мин", u"Меридиональное сечение"),
    "sag": (1, r"$M'$, мм",   r"$\Delta x'$, мм",   r"$\\psi'$, мин", u"Сагиттальное сечение"),
    "cfs": (1, r"$-\omega'$, град", r"$\Delta s'$, мм",   r"$\Delta s'$, дптр", u"Хроматизм положения"),
    "dist":(1, r"$-\omega'$, град", r"$\Delta y'$, мм",   r"$\Delta\omega', \%", u"Дисторсия"),
    "ast": (1, r"$-\omega'$, град", r"$L'$, мм",r"$L'cos\omega'$, дптр", u"Астигматизм"),
}

preamble=[
    '\\documentclass{article}\n',
    '\\usepackage{tikz}\n',
    '\\usepackage{pgfplots}\n',
    '\\usepackage[utf8]{inputenc}\n',
    '\\usepackage[english,russian]{babel}\n',
    '\\begin{document}\n',
    '\\begin{tikzpicture}\n',
    '\t\\begin{axis}[\n',
    '\t\taxis lines=middle,\n',
    '\t\tno markers,\n',
    '\t\tscaled x ticks=false,\n',
    '\t\ttick label style={\n',
    '\t\t\t/pgf/number format/fixed,\n',
    '\t\t\t/pgf/number format/precision=3,\n',
    '\t\t\t/pgf/number format/use comma,\n',
    '\t\t},\n',
    '\t\tenlarge y limits={rel=0.2},\n',
    '\t\tenlarge x limits={rel=0.2},\n',
    '\t\txlabel style={at=(current axis.right of origin), anchor=west},\n',
    '\t\tylabel style={at=(current axis.above origin), anchor=south},\n',
]
postfix=[
    '\t\end{axis};\n',
    '\\end{tikzpicture}\n',
    '\\end{document}\n'
]

spectral_lines=[
    (759.37,"А"),
    (686.719,"B"),
    (656.281,"C"),
    (627.661,"а"),
    (589.592,"D1"),
    (588.995,"D2"),
    (587.5618,"D3"),
    (546.073,"е"),
    (527.039,"E2"),
    (518.362,"b1"),
    (517.27,"b2"),
    (516.891,"b3"),
    (516.733,"b4"),
    (495.761,"c"),
    (486.134,"F"),
    (466.814,"d"),
    (438.355,"e'"),
    (434.047,"G'"),
    (430.79,"G"),
    (430.774,"g"),
    (410.175,"h"),
    (396.847,"H"),
]

ap = argparse.ArgumentParser(
    description="Create files for TikZ plotting from Zemax text report files."
)
ap.add_argument(
    'path', 
    metavar='path', 
    type=str,
    help='path to zmx report file'
)
ap.add_argument(
    'outputdir', 
    metavar='outputdir', 
    type=str,
    help='directory for CSV and TeX file output'
)
ap.add_argument(
    "-t",
    action='store', 
    dest='type', 
    required=True, 
    type=str, 
    choices=["tang", "sag", "lon", "cfs", "ast","sph"], 
    help='type of plot'
)
ap.add_argument(
    "--afocal", 
    "-a",
    required=False, 
    help="system is afocal", 
    action="store_true", 
    dest='afocal'
)
ap.add_argument(
    "--yscale", 
    dest='yscale', 
    type=float,
    required=False, 
    default=1,
    help="y coord scale multiplier (for wide beam aberrations equals exit pupil radius), default 1"
)
ap.add_argument(
    "--enc",
    required=False,
    type=str,
    dest="enc",
    choices=["ansi","utf-8","utf-16-le","utf-16-be","utf-32"],
    default="ansi",
    help="input file encoding",
)
args=ap.parse_args()

plottype = args.type
x_column_index, y_label, x_label, x_label_afocal, label = plot_types.get(plottype)
if args.afocal:
    x_label = x_label_afocal
fpath = os.path.normpath(args.path)
fname = os.path.splitext(os.path.basename(fpath))[0]
outdir = os.path.normpath(args.outputdir)
if not os.path.isdir(outdir):
    print('ERROR: output directory specified with --outputdir key not recognized as valid')
    os._exit(1)

report = open(fpath, 'r', encoding=args.enc)
is_sagittal=False
waves=[]
coords=[]
for line in report:
    words=line.split()
    if len(words)==0:
        continue
    if is_number(words[0]):
        if not ((plottype=="sag" and not is_sagittal) or (plottype=="tang" and is_sagittal)):
            x=[] #list of x coords in case there are more than one
            y=abs(float(words[0]))
            if plottype=='lon':
                for i,wvl in enumerate(waves):
                    x.append(float(words[x_column_index+i]))
            elif plottype=='ast':
                x.append(float(words[x_column_index]))
                x.append(float(words[x_column_index+1]))
            else:
                x=float(words[x_column_index])
            coords.append([y,x])
    else:
        if line.find("Sagittal fan") > -1:
            is_sagittal=True
        if plottype=='lon' and line.find("Rel. Pupil") > -1:
            #generate wavelength list
            waves=[float(word) for word in words if is_number(word)]
report.close()
if len(coords) == 0:
    print("ERROR: no coordinates found in file")
    print("Probably wrong input file encoding. Either use the --encoding option to specify the right one or convert to ASCII")

csv_files=[]
for col in range(0,len(coords[0][1])):
    #for each x var we'll create a table
    table=[]
    for coord in coords:
        x=coord[1][col]
        y=coord[0]
        table.append((x,y))
    fname_csv=f'{fname.replace("_","")}wave{col}.csv'
    csv_files.append(fname_csv)
    fpath_csv=os.path.join(outdir,fname_csv)
    with open(fpath_csv, "w") as csvfile:
        for row in table:
            csvfile.write(f"{row[0]},{row[1]}\n")
print(f"Saving {len(csv_files)} CSV table(s) of coordinates in /{outdir}/.")

plot_labels=[]
if plottype=='lon':
    for wave in waves:
        waveletter=''
        for spectral_line in spectral_lines:
            if abs(spectral_line[0]-wave*1000) < 10.0:
                waveletter=spectral_line[1]
        texcode = '$\lambda_{' + waveletter + '}$' 
        plot_labels.append(texcode)
elif plottype=='ast':
    if args.afocal:
        plot_labels.append(r"$L'_m cos\omega$")
        plot_labels.append(r"$L'_s cos\omega$")
    else:
        plot_labels.append("$L'_m$")
        plot_labels.append("$L'_s$")

ytick='0.7,1'
yticklabels=f'{round(args.yscale*0.7,1)},{args.yscale}'
#these are modified only for pupil coords
#for field and wavelength ticks are extracted from CSV

fname_tex = fname.replace("_","")+".tex"
fpath_tex = os.path.join(outdir, fname_tex)
with open(fpath_tex,'w', encoding='utf-8') as texfile:
    texfile.writelines(preamble)
    #texfile.write('\t\ttitle={' + label + '},\n')
    texfile.write('\t\txlabel={' + x_label + '},\n')
    texfile.write('\t\tylabel={' + y_label + '},\n')
    if not (plottype=='ast' or plottype=='dist' or plottype=='cfs'):
        texfile.write('\t\tytick={'+ytick+'},\n')
        texfile.write('\t\tyticklabels={'+yticklabels+'},\n')
    texfile.write('\t\t]\n')
    for i,csv_name in enumerate(csv_files):
        label=''
        if plottype=='lon' or plottype=='ast':
            label= ' node[above]{' + plot_labels[i] + '}'
        texfile.write('\t\t\\addplot[color=black] table[col sep=comma]{' + csv_name + '}' + label + ';\n')
    texfile.writelines(postfix)
print("Saving TeX file.")