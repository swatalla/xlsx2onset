function ofe = offset_end(i, sgTemp, tStim)

while sgTemp(i) > tStim(1)
    i = i+1;
end

ofe = i;