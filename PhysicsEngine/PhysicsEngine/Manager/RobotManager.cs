using PhysicsEngine.Model;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace PhysicsEngine.Manager
{
    [Guid("79629CB0-947F-4A71-BE5C-5C53FF66D298")]
    public interface IRobotManager
    {
        Vector GetRobotPosition(int n);

        void SetRobotPosition(int n, ref Vector vector);
    }

    [Guid("44B84867-82CE-4B6D-9ACF-299EADACF2A0"), ClassInterface(ClassInterfaceType.None)]
    public class RobotManager : IRobotManager
    {
        private static readonly Dictionary<int, Robot> _robots = new Dictionary<int, Robot>();

        public Vector GetRobotPosition(int n)
        {
            if (!_robots.ContainsKey(n))
                _robots.Add(n, new Robot());

            return _robots[n].Position;
        }

        public void SetRobotPosition(int n, ref Vector position)
        {
            if (!_robots.ContainsKey(n))
                _robots.Add(n, new Robot());

            _robots[n].Position = position;
        }
    }
}