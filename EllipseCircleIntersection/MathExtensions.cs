using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using devDept.Geometry;

namespace EllipseCircleIntersection
{
    public static class MathExtensions
    {
        public static double Pow2(this double n)
        {
            return Math.Pow(n, 2);
        }

        public static double Pow3(this double n)
        {
            return Math.Pow(n, 3);
        }

        public static double Pow4(this double n)
        {
            return Math.Pow(n, 4);
        }

        public static Tuple<double, double> GetSolutionPoint(double z, double r, double xoK, double yoK)
        {
            var zPow2 = z.Pow2();
            var x = xoK + r * (1 - zPow2) / (1.0 + zPow2);
            var y = yoK + r * (2 * z) / (1 + zPow2);
            return Tuple.Create( x, y);
        }

        public static Point2D[] GetPoints(double xoK, double yoK, double radius, double xoE, double yoE, double ellA, double ellB)
        {
            var aInit = ellA.Pow2() * ((yoK - yoE).Pow2() - ellB.Pow2()) + ellB.Pow2() * ((xoK - xoE) - radius).Pow2();
            var bInit = 4 * ellA.Pow2() * radius * (yoK - yoE);
            var cInit = 2 * (ellA.Pow2() * ((yoK - yoE).Pow2() - ellB.Pow2() + 2 * radius.Pow2()) + ellB.Pow2() * ((xoK - xoE).Pow2() - radius.Pow2()));
            var dInit = 4 * ellA.Pow2() * radius * (yoK - yoE);
            var eInit = ellA.Pow2() * ((yoK - yoE).Pow2() - ellB.Pow2()) + ellB.Pow2() * ((xoK - xoE) + radius).Pow2();

            var a = bInit / aInit;
            var b = cInit / aInit;
            var c = dInit / aInit;
            var d = eInit / aInit;

            var p = b - 3 * a.Pow2() / 8;
            var q = a.Pow3() / 8 - a * b / 2 + c;
            var r = -(3 * a.Pow4() - 16 * a.Pow2() * b + 64 * a * c - 256 * d) / 256.0;

            var m = -2 * p;
            var n = p.Pow2() - 4 * r;
            var o = q.Pow2();

            var pN = n - m.Pow2() / 3;
            var qN = 2 * m.Pow3() / 27 - m * n / 3 + o;

            var R = (qN / 2).Pow2() + (pN / 3).Pow3();

            var tolerance = 1e-7;
            List<double> realSolutions;

            //Case2: Can be reduced to quadratic equation
            if (Math.Abs(a) < tolerance && Math.Abs(c) < tolerance)
            {
                var x1 = -b / 2 + Math.Sqrt((-b / 2).Pow2() - d);
                var x2 = -b / 2 - Math.Sqrt((-b / 2).Pow2() - d);

                var z1 = Math.Sqrt(x1);
                var z2 = Math.Sqrt(x2);
                var z3 = -z1;
                var z4 = -z2;

                realSolutions = new List<double>() { z1, z2, z3, z4 };
            }
            else
            {
                Complex y1;
                Complex y2;
                Complex y3;
                if (R >= 0)
                {
                    var T = Math.Sqrt(R);
                    var c1 = -qN / 2 + T;
                    var c2 = -qN / 2 - T;
                    var u = Math.Sign(c1) * Complex.Pow(c1, 1.0 / 3).Magnitude;
                    var v = Math.Sign(c2) * Complex.Pow(c2, 1.0 / 3).Magnitude;
                    y1 = u + v;
                    y2 = -(u + v) / 2 - ((u - v) / 2) * Math.Sqrt(3) * Complex.ImaginaryOne;
                    y3 = -(u + v) / 2 + ((u - v) / 2) * Math.Sqrt(3) * Complex.ImaginaryOne;
                }
                else
                {
                    var u = Math.Sqrt(-(pN / 3).Pow3());
                    var w = Math.Acos(-qN / (2 * u));
                    y1 = 2 * Math.Pow(u, 1.0 / 3) * Math.Cos(w / 3);
                    y2 = 2 * Math.Pow(u, 1.0 / 3) * Math.Cos(w / 3 + 2 * Math.PI / 3);
                    y3 = 2 * Math.Pow(u, 1.0 / 3) * Math.Cos(w / 3 + 4 * Math.PI / 3);
                }

                var x1 = y1 - (m / 3);
                var x2 = y2 - (m / 3);
                var x3 = y3 - (m / 3);

                var sqrt1 = Complex.Sqrt(-x1);
                var sqrt2 = Complex.Sqrt(-x2);
                var sqrt3 = Complex.Sqrt(-x3);

                var sum1 = sqrt1 * sqrt2 * sqrt3;
                if (!(Math.Abs(sum1.Real + q) <= 1e-5))
                {
                    sqrt1 = -sqrt1;
                    sqrt2 = -sqrt2;
                    sqrt3 = -sqrt3;
                }

                var z1 = (sqrt1 + sqrt2 + sqrt3) / 2;
                var z2 = (sqrt1 - sqrt2 - sqrt3) / 2;
                var z3 = (-sqrt1 + sqrt2 - sqrt3) / 2;
                var z4 = (-sqrt1 - sqrt2 + sqrt3) / 2;

                z1 = z1 - a / 4;
                z2 = z2 - a / 4;
                z3 = z3 - a / 4;
                z4 = z4 - a / 4;

                realSolutions = new[] { z1, z2, z3, z4 }
                    .Where(pt => Math.Abs(pt.Imaginary) < 1e-4)
                    .Select(sol => sol.Real)
                    .ToList();

            }

            var solutionPoints = realSolutions
                .Select(x => GetSolutionPoint(x, radius, xoK, yoK))
                .Distinct()
                .Select(w => new Point2D(w.Item1,w.Item2))
                .ToArray();
            return solutionPoints;
        }
    }
}
