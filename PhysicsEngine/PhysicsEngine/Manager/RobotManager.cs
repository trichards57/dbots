using PhysicsEngine.Model;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace PhysicsEngine.Manager
{
    [Guid("79629CB0-947F-4A71-BE5C-5C53FF66D298")]
    public interface IRobotManager
    {
        Vector GetActualVelocity(int n);

        bool GetExists(int n);

        float GetRadius(int n);

        Vector GetRobotPosition(int n);

        Vector GetVelocity(int n);

        void SetActualVelocity(int n, ref Vector vector);

        void SetExists(int n, bool value);

        void SetRadius(int n, float radius);

        void SetRobotPosition(int n, ref Vector vector);

        void SetVelocity(int n, ref Vector vector);
    }

    [Guid("44B84867-82CE-4B6D-9ACF-299EADACF2A0"), ClassInterface(ClassInterfaceType.None)]
    public class RobotManager : IRobotManager
    {
        internal static Dictionary<int, Robot> Robots { get; } = new Dictionary<int, Robot>();

        public Vector GetActualVelocity(int n)
        {
            if (!Robots.ContainsKey(n))
                Robots.Add(n, new Robot());

            return Robots[n].ActualVelocity;
        }

        public bool GetExists(int n)
        {
            if (!Robots.ContainsKey(n))
                Robots.Add(n, new Robot());

            return Robots[n].Exists;
        }

        public float GetRadius(int n)
        {
            if (!Robots.ContainsKey(n))
                Robots.Add(n, new Robot());

            return Robots[n].Radius;
        }

        public Vector GetRobotPosition(int n)
        {
            if (!Robots.ContainsKey(n))
                Robots.Add(n, new Robot());

            return Robots[n].Position;
        }

        public Vector GetVelocity(int n)
        {
            if (!Robots.ContainsKey(n))
                Robots.Add(n, new Robot());

            return Robots[n].Velocity;
        }

        public void SetActualVelocity(int n, ref Vector vector)
        {
            if (!Robots.ContainsKey(n))
                Robots.Add(n, new Robot());

            Robots[n].ActualVelocity = vector;
        }

        public void SetExists(int n, bool value)
        {
            if (!Robots.ContainsKey(n))
                Robots.Add(n, new Robot());

            Robots[n].Exists = value;
        }

        public void SetRadius(int n, float radius)
        {
            if (!Robots.ContainsKey(n))
                Robots.Add(n, new Robot());

            Robots[n].Radius = radius;
        }

        public void SetRobotPosition(int n, ref Vector position)
        {
            if (!Robots.ContainsKey(n))
                Robots.Add(n, new Robot());

            Robots[n].Position = position;
        }

        public void SetVelocity(int n, ref Vector vector)
        {
            if (!Robots.ContainsKey(n))
                Robots.Add(n, new Robot());

            Robots[n].Velocity = vector;
        }
    }
}