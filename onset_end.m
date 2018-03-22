function onse = onset_end(i, sgTemp, tStim)

j = i;
if sgTemp(i+1) >= sgTemp(i)
    if sgTemp(i+1) <= tStim(2) && sgTemp(i+1) >= tStim(1)
        while sgTemp(i+1) <= tStim(2) && sgTemp(i+1) > tStim(1)
            i = i+1;
        end
    end
end

if ~isempty(find(sgTemp(j:i) == tStim(2),1))
    i = find(sgTemp(j:i) == tStim(2),1) + j;
end
    h = i;
    g = i;

while sgTemp(g) > tStim(1)
    g = g+1;
end

mda = max(sgTemp(i:g));

if mda > tStim(2)
    if sgTemp(i+1) <= tStim(3) && sgTemp(i+1) >= tStim(2)
        while sgTemp(i+1) <= tStim(3)...
                && sgTemp(i+1) < tStim(4)...
                && sgTemp(i+1) >= tStim(2)
            i = i+1;
        end
        if ~isempty(find(sgTemp(h:i) == tStim(3),1) + h)
            i = find(sgTemp(h:i) == tStim(3),1) + h;
        end
    end
    
    f = i;
    
    while sgTemp(f) > tStim(2)
        f = f+1;
    end
    
    mdb = max(sgTemp(i:f));
    
    if mdb > tStim(3)
        if sgTemp(i+1) <= tStim(4) && sgTemp(i+1) >= tStim(3)
            while sgTemp(i) < tStim(4)
                i = i+1;
            end
            i = i+1;
        end
    end
end


onse = i-1;

end