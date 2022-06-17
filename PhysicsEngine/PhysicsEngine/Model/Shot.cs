using System.Runtime.InteropServices;

namespace PhysicsEngine.Model
{
    public struct Shot
    {
        public short Age;

        public int Color;

        [MarshalAs(UnmanagedType.SafeArray)]
        public Block[] Dna;

        public int DnaLength;

        public float Energy;

        [MarshalAs(UnmanagedType.Bool)]
        public bool Exists;

        [MarshalAs(UnmanagedType.Bool)]
        public bool Flash;

        [MarshalAs(UnmanagedType.BStr)]
        public string FromSpecies;

        [MarshalAs(UnmanagedType.Bool)]
        public bool FromVeg;

        public short GeneNumber;
        public short MemoryLocation;

        public short MemoryValue;

        public Vector OldPosition;

        public short Parent;

        public Vector Position;

        public float Range;

        public short ShotType;

        [MarshalAs(UnmanagedType.Bool)]
        public bool Stored;

        public short Type;
        public float Value;
        public Vector Velocity;
    }
}