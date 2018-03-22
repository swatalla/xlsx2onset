function ofs = offset_start(i, sgTemp, ~)

while sgTemp(i) <= sgTemp(i+1)
    i = i + 1;
end

ofs = i;