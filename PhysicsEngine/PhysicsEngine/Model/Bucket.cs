using System.Collections.Generic;

namespace PhysicsEngine.Model
{
    public struct Bucket
    {
        public List<Point> AdjacentBuckets { get; set; }
        public HashSet<int> RobotsIds { get; set; }
    }
}