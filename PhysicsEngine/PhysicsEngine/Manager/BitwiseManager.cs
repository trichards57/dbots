using System.Runtime.InteropServices;

namespace PhysicsEngine.Manager
{
    [Guid("CBF82E12-9803-4F9E-A4C3-6911D88F7CB4")]
    public interface IBitwiseManager
    {
        int And(int num1, int num2);

        int Decrement(int num);

        int Increment(int num);

        int Invert(int num);

        int Or(int num1, int num2);

        int ShiftLeft(int num);

        int ShitRight(int num);

        int Xor(int num1, int num2);
    }

    [Guid("95A6769F-6C3D-4B01-9F41-B189F9EC83F2"), ClassInterface(ClassInterfaceType.None)]
    public class BitwiseManager : IBitwiseManager
    {
        public int And(int num1, int num2) => num1 & num2;

        public int Decrement(int num) => num - 1;

        public int Increment(int num) => num + 1;

        public int Invert(int num) => ~num;

        public int Or(int num1, int num2) => num1 | num2;

        public int ShiftLeft(int num) => num << 1;

        public int ShitRight(int num) => num >> 1;

        public int Xor(int num1, int num2) => num1 ^ num2;
    }
}