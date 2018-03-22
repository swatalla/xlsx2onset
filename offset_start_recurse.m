function ofs = offset_start_recurse(i, sgTemp, tStim)
while sgTemp(i) == tStim(1)
    i = i+1;
end
i = i-1;
if sgTemp(i+1) >= sgTemp(i)
    if sgTemp(i+1) <= tStim(2) && sgTemp(i+1) > tStim(1)
        while sgTemp(i+1) <= tStim(2) && sgTemp(i+1) < tStim(3)
            i = i+1;
        end
    end
    if sgTemp(i+1) <= tStim(3) && sgTemp(i+1) > tStim(2)
        while sgTemp(i+1) <= tStim(3) && sgTemp(i+1) < tStim(4)
            i = i+1;
        end
        if sgTemp(i+1) <= tStim(4) && sgTemp(i+1) > tStim(3)
            while sgTemp(i) < tStim(4)
                i = i+1;
            end
            if sgTemp(i+1) >= tStim(4)
                while sgTemp(i) >= tStim(4)
                    i = i+1;
                end
            end
        end
    end
end

ofs=i-1;

end