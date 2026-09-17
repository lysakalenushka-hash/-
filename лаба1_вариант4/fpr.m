function F=fpr(t,h)
global ro pn p ak hg g v s
p(10) = pn * hg(1) / (hg(1) - h(1));
p(7)  = p(10) + ro * g * h(1);
p(11) = pn * hg(2) / (hg(2) - h(2));
p(8)  = p(11) + ro * g * h(2);
p(12) = pn * hg(3) / (hg(3) - h(3));
p(9)  = p(12) + ro * g * h(3);
v(1) = ak(1) * sign(p(1) - p(8)) * sqrt(abs(p(1) - p(8)));
v(2) = ak(2) * sign(p(2) - p(9)) * sqrt(abs(p(2) - p(9)));
v(3) = ak(3) * sign(p(7) - p(8)) * sqrt(abs(p(7) - p(8)));
v(4) = ak(4) * sign(p(8) - p(9)) * sqrt(abs(p(8) - p(9)));
v(5) = ak(5) * sign(p(7) - p(3)) * sqrt(abs(p(7) - p(3)));
v(6) = ak(6) * sign(p(8) - p(4)) * sqrt(abs(p(8) - p(4)));
v(7) = ak(7) * sign(p(8) - p(5)) * sqrt(abs(p(8) - p(5)));
v(8) = ak(8) * sign(p(9) - p(6)) * sqrt(abs(p(9) - p(6)));
F=[(-v(3)-v(5))/s(1);
   (v(1)+v(3)-v(4)-v(6)-v(7))/s(2);
   (v(2)+v(4)-v(8))/s(3)];
end
