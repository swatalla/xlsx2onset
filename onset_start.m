function ons = onset_start(i, sgTemp, tStim)

while sgTemp(i) <= tStim(1)
    i = i+1;
end

ons = i-1;