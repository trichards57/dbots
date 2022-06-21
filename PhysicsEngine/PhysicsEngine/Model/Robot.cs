namespace PhysicsEngine.Model
{
    public class Robot
    {
        public Vector ActualVelocity { get; set; }
        public bool Exists { get; set; }
        public Vector Position { get; set; }
        public float Radius { get; set; }
        public Vector Velocity { get; set; }
    }
}