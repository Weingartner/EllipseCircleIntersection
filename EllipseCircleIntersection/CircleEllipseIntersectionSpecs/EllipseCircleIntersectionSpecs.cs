using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using devDept.Eyeshot.Entities;
using devDept.Geometry;
using EllipseCircleIntersection;
using FluentAssertions;
using Xunit;

namespace CircleEllipseIntersectionSpecs
{
    public class EllipseCircleIntersectionSpecs
    {
        public static IEnumerable<object[]> EllipseCircleTestData
        {
            get
            {
                yield return new object[] { 1, 2, 3, 4, 5, 6, 7 };
                yield return new object[] { 0, 0, 3, 4, 5, 6, 7 };
                yield return new object[] { 0, 0, 1, 4, 5, 6, 7 };
                yield return new object[] { 3, 0, 1, 4, 5, 2, 7 };
                yield return new object[] { 1, 1, 3, 1, 1, 2, 7 };
                yield return new object[] { 5, 0, 1, 4, 5, 2, 7 };
                yield return new object[] { 0, 6, 3, 0, 1, 2, 7 };
                yield return new object[] { 0, 5, 3, 0, 1, 2, 7 };
                yield return new object[] { 0, 10, 3, 0, 1, 2, 7 };
                yield return new object[] { 0, 11, 3, 0, 1, 2, 7 };
                yield return new object[] { 0, 12, 3, 0, 1, 2, 7 };
            }
        }

        /// <summary>
        /// https://www.arndt-bruenner.de/mathe/scripts/polynome.htm
        /// https://math.stackexchange.com/questions/3419984/find-the-intersection-of-a-circle-and-an-ellipse
        /// http://www.mathe.tu-freiberg.de/~hebisch/cafe/viertergrad.pdf
        /// </summary>
        /// <param name="xoK"></param>
        /// <param name="yoK"></param>
        /// <param name="radius"></param>
        /// <param name="xoE"></param>
        /// <param name="yoE"></param>
        /// <param name="ellA"></param>
        /// <param name="ellB"></param>
        [Theory]
        [MemberData(nameof(EllipseCircleTestData))]
        public void AnalyticEllipseCircleIntersectionForErwinSpec(double xoK, double yoK, double radius, double xoE, double yoE, double ellA, double ellB)
        {

            var circle = new Circle(new Point3D(xoK, yoK), radius);
            var ellipse = new Ellipse(new Point3D(xoE, yoE), ellA, ellB);

            MathExtensions.GetPoints(xoK,yoK,radius,xoE,yoE,ellA,ellB)
                .ToList()
                .ForEach(pt =>
                {
                    var point3D = new Point3D(pt.X,pt.Y);
                    circle.ClosestPointTo(point3D , out var t);
                    circle.PointAt(t).DistanceTo(point3D).Should().BeLessOrEqualTo(1e-7);

                    ellipse.ClosestPointTo(point3D , out var rt);
                    ellipse.PointAt(rt).DistanceTo(point3D).Should().BeLessOrEqualTo(1e-7);
                });

            //new Assembly3D()
            //    .Add(circle, Color.Red, 2)
            //    .Add(ellipse, Color.Blue, 2)
            //    .AddCloud(solutionPoints, Color.Orange, 10)
            //    .Log($"R: {R}, a: {aInit}, b: {bInit}, c: {cInit}, d: {dInit}, e: {eInit}")
            //    .Wait();
        }
    }
}
