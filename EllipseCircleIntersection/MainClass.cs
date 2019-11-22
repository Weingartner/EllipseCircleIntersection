using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace EllipseCircleIntersection
{
    /// <summary>
    /// https://www.codeproject.com/Articles/3511/Exposing-NET-Components-to-COM#vb
    /// </summary>
	[Guid("D6F88E95-8A27-4ae6-B6DE-0542A0FC7039")]
	[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
	public interface IEllCircInt
	{
		[DispId(1)]
		double[] GetIntersectionPts(double xoK, double yoK, double radius, double xoE, double yoE, double ellA, double ellB);
	}

	[Guid("13FE32AD-4BF8-495f-AB4D-6C61BD463EA4")]
	[ClassInterface(ClassInterfaceType.None)]
	[ProgId("Tester.Numbers")]
	public class EllCircInt : IEllCircInt
	{
		public EllCircInt(){}

        /// <summary>
        /// Calculates the intersection points between an ellipse and a circle in a general position.
        /// </summary>
        /// <param name="xoK"></param>
        /// <param name="yoK"></param>
        /// <param name="radius"></param>
        /// <param name="xoE"></param>
        /// <param name="yoE"></param>
        /// <param name="ellA"></param>
        /// <param name="ellB"></param>
        /// <returns></returns>
		public double[] GetIntersectionPts(double xoK, double yoK, double radius, double xoE, double yoE, double ellA, double ellB)
		{
			return MathExtensions.GetPoints(xoK, yoK, radius, xoE, yoE, ellA, ellB)
                .SelectMany(pt=>new List<double> {pt.X,pt.Y})
                .ToArray();
        }
    }
}
