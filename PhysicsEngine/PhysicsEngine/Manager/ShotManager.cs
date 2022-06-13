using PhysicsEngine.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PhysicsEngine.Manager
{
    [Guid("0E1EEC48-FF16-4E41-B4D8-FEB855A5C019")]
    public interface IShotManager
    {
        [DispId(2)]
        Shot GetShot(int i);

        [DispId(1)]
        int NewShot(short parentBot, short shotType, float value, float startEnergy, float startRange,
            [MarshalAs(UnmanagedType.BStr)] string fromSpecies,
            [MarshalAs(UnmanagedType.Bool)] bool fromVeg,
            int color, short memoryLocation, short memoryValue, float shotAim, Vector startPosition,
            Vector startVelocity,
            [MarshalAs(UnmanagedType.Bool)] bool offset = false);

        [DispId(3)]
        int CreateShot();
    }

    [Guid("E514AE43-3D40-48A3-9C62-D54058826EAA"), ClassInterface(ClassInterfaceType.None)]
    public class ShotManager : IShotManager
    {
        private readonly Dictionary<int, Shot> _shots = new Dictionary<int, Shot>();

        public Shot GetShot(int i)
        {
            if (_shots.ContainsKey(i))
                return _shots[i];
            return new Shot();
        }

        public int CreateShot()
        {
            return FirstSlot();
        }

        public int NewShot(short parentBot, short shotType, float value, float startEnergy, float startRange,
                    [MarshalAs(UnmanagedType.BStr)] string fromSpecies,
            [MarshalAs(UnmanagedType.Bool)] bool fromVeg,
            int color, short memoryLocation, short memoryValue, float shotAim, Vector startPosition,
            Vector startVelocity,
            [MarshalAs(UnmanagedType.Bool)] bool offset = false)
        {
            var i = FirstSlot();

            if (i == -1)
                return -1; // Ran out of spaces for shots, return an invalid ID

            var shot = new Shot
            {
                Exists = true,
                Age = 0,
                Parent = parentBot,
                FromSpecies = fromSpecies,
                FromVeg = fromVeg,
                Color = color,
                Value = (float)Math.Round(value)
            };

            if (shotType > 0 || shotType == -100)
                shot.Type = shotType;
            else
            {
                shot.Type = (short)-(Math.Abs(shotType) % 8);
                if (shot.Type == 0) shot.Type = -8;
            }

            if (shot.Type == -2)
                shot.Color = 0xFFFFFF;

            shot.MemoryLocation = memoryLocation;
            shot.MemoryValue = memoryValue;
            shot.Position = startPosition;
            shot.Velocity = startVelocity;
            shot.OldPosition = startPosition - startVelocity;
            shot.Energy = startEnergy;
            shot.Range = startRange;

            _shots[i] = shot;

            return i;
        }

        private int FirstSlot()
        {
            for (var i = 0; i < int.MaxValue; i++)
            {
                if (!_shots.ContainsKey(i))
                    return i;
            }

            return -1;
        }
    }
}