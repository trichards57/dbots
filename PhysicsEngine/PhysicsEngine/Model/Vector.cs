using System;
using System.Numerics;

namespace PhysicsEngine.Model
{
    public struct Vector
    {
        public float X;
        public float Y;

        public static implicit operator Vector2(Vector v)
        {
            return new Vector2(v.X, v.Y);
        }

        public static Vector operator -(Vector v1, Vector v2)
        {
            return new Vector { X = v1.X - v2.X, Y = v1.Y - v2.Y };
        }

        public static Vector operator *(Vector v1, float k)
        {
            var x = v1.X * k;
            var y = v1.Y * k;

            if (Math.Abs(x) > 32000)
                x = Math.Sign(x) * 32000;
            if (Math.Abs(y) > 32000)
                y = Math.Sign(y) * 32000;

            return new Vector { X = x, Y = y };
        }

        public static Vector operator *(Vector v1, double k)
        {
            var x = v1.X * k;
            var y = v1.Y * k;

            if (Math.Abs(x) > 32000)
                x = Math.Sign(x) * 32000;
            if (Math.Abs(y) > 32000)
                y = Math.Sign(y) * 32000;

            return new Vector { X = (float)x, Y = (float)y };
        }

        public static Vector operator +(Vector v1, Vector v2)
        {
            return new Vector { X = v1.X + v2.X, Y = v1.Y + v2.Y };
        }

        public float MagnitudeSquared() => X * X + Y * Y;
    }
}