using PhysicsEngine.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace PhysicsEngine.Manager
{
    [Guid("0E1EEC48-FF16-4E41-B4D8-FEB855A5C019")]
    public interface IShotManager
    {
        CollisionReport CheckForCollision(float maxBotShotSeperation, int shot);

        void Clear();

        int CreateShot();

        void DeleteShot(int a);

        int GetMaxShot();

        Shot GetShot(int i);

        void LoadShots(string file);

        void SaveShots(string file);

        void SetShot(int i, ref Shot value);

        UpdateResult UpdateShotsCollisions(float maxBotShotSeperation, float minBotRadius, bool upDnConnected, bool dxSxConnected, ref Model.Vector fieldSize);

        void UpdateShotsPosition(bool noShotDecay, bool noWasteShowDecay);
    }

    public struct CollisionReport
    {
        public int Bot;
        public int Shot;
    }

    public struct UpdateResult
    {
        [MarshalAs(UnmanagedType.SafeArray)]
        public CollisionReport[] Collisions;

        public int NumShots;
        public float TotalEnergy;
    }

    [Guid("E514AE43-3D40-48A3-9C62-D54058826EAA"), ClassInterface(ClassInterfaceType.None)]
    public class ShotManager : IShotManager
    {
        private static readonly Dictionary<int, Shot> _shots = new Dictionary<int, Shot>();

        public CollisionReport CheckForCollision(float maxBotShotSeperation, int shot)
        {
            var sh = _shots[shot];

            var collisions = RobotManager.Robots
                .AsParallel()
                .Where(r => r.Value.Exists
                && r.Key != sh.Parent
                && Math.Abs(sh.OldPosition.X - r.Value.Position.X) < maxBotShotSeperation
                && Math.Abs(sh.OldPosition.Y - r.Value.Position.Y) < maxBotShotSeperation)
                .Select(rob =>
               {
                   double hitTime;
                   var robPosition = rob.Value.Position - rob.Value.Velocity + rob.Value.ActualVelocity;

                   var vectorToBot = sh.Position - robPosition;
                   var distanceToBotSquared = vectorToBot.MagnitudeSquared();
                   var radiusSquared = Math.Pow(rob.Value.Radius, 2);

                   if (distanceToBotSquared < radiusSquared)
                   {
                       return new Collision(0, (short)rob.Key);
                   }

                   var relativeVelocity = sh.Velocity - rob.Value.ActualVelocity;
                   var relativeSpeedSquared = relativeVelocity.MagnitudeSquared();
                   if (relativeSpeedSquared == 0)
                       return null as Collision?;

                   var dDotP = Vector2.Dot(relativeVelocity, vectorToBot);
                   var x = -dDotP;
                   var y = (dDotP * dDotP) - relativeSpeedSquared * (distanceToBotSquared - radiusSquared);

                   if (y < 0)
                       return null as Collision?;

                   y = (float)Math.Sqrt(y);

                   var time0 = (x - y) / relativeSpeedSquared;
                   var time1 = (x + y) / relativeSpeedSquared;

                   var useTime0 = false;
                   var useTime1 = false;

                   if (!(time0 <= 0 || time0 >= 1))
                       useTime0 = true;
                   if (!(time1 <= 0 || time1 >= 1))
                       useTime1 = true;

                   if (!useTime1 && !useTime1)
                       return null;

                   if (useTime0 && useTime1)
                   {
                       hitTime = Math.Min(time0, time1);
                   }
                   else if (useTime0)
                   {
                       hitTime = time0;
                   }
                   else
                   {
                       hitTime = time1;
                   }
                   return new Collision(hitTime, (short)rob.Key);
               });

            var earliestCollision = collisions.Where(c => c.HasValue).OrderBy(c => c.Value.Time).FirstOrDefault();

            if (earliestCollision.HasValue)
            {
                sh.Position = sh.Velocity * earliestCollision.Value.Time + sh.Position;
                _shots[shot] = sh;
            }

            return new CollisionReport
            {
                Bot = earliestCollision?.Robot ?? 0,
                Shot = shot
            };
        }

        public void Clear()
        {
            _shots.Clear();
        }

        public int CreateShot()
        {
            return FirstSlot();
        }

        public void DeleteShot(int a)
        {
            _shots.Remove(a);
        }

        public int GetMaxShot()
        {
            return _shots.Count > 0 ? _shots.Select(kvp => kvp.Key).Max() : 1;
        }

        public Shot GetShot(int i)
        {
            Shot? res = null;

            if (_shots.ContainsKey(i))
                res = _shots[i];

            if (res == null)
            {
                res = new Shot();
                _shots[i] = res.Value;
            }

            return res.Value;
        }

        public void LoadShots(string file)
        {
            try
            {
                var newShots = JsonSerializer.Deserialize<Dictionary<int, Shot>>(File.ReadAllText(file));

                _shots.Clear();

                foreach (var kvp in newShots)
                    _shots.Add(kvp.Key, kvp.Value);
            }
            catch (IOException ex)
            {
                Debug.WriteLine($"Could not read shots : {ex}");
            }
        }

        public void SaveShots(string file)
        {
            try
            {
                ClearRemoved();
                File.WriteAllText(file, JsonSerializer.Serialize(_shots, new JsonSerializerOptions { IncludeFields = true }));
            }
            catch (IOException ex)
            {
                Debug.WriteLine($"Could not write shots : {ex}");
            }
        }

        public void SetShot(int i, ref Shot value)
        {
            if (value.Exists)
                _shots[i] = value;
            else
                _shots.Remove(i);
        }

        public UpdateResult UpdateShotsCollisions(float maxBotShotSeperation, float minBotRadius, bool upDnConnected, bool dxSxConnected, ref Model.Vector fieldSize)
        {
            var keys = _shots.Where(s => s.Value.Flash || !s.Value.Exists).Select(s => s.Key).ToList();

            foreach (var key in keys)
                _shots.Remove(key);

            keys = _shots.Keys.ToList();

            var collisions = new List<CollisionReport>();

            foreach (var key in keys)
            {
                _shots[key] = FieldBorderCollision(_shots[key], upDnConnected, dxSxConnected, fieldSize);
                if (_shots[key].Type != -100 && !_shots[key].Stored)
                {
                    var res = CheckForCollision(maxBotShotSeperation, key);
                    if (res.Bot != 0)
                        collisions.Add(res);
                }
            }

            var result = new UpdateResult
            {
                NumShots = _shots.Count,
                TotalEnergy = _shots.Where(s => s.Value.Type == -2).Sum(s => s.Value.Energy),
                Collisions = collisions.ToArray()
            };

            return result;
        }

        public void UpdateShotsPosition(bool noShotDecay, bool noWasteShowDecay)
        {
            var keys = _shots.Keys.ToList();

            foreach (var key in keys)
            {
                var sh = _shots[key];
                sh.OldPosition = sh.Position;
                sh.Position += sh.Velocity;

                if ((!noShotDecay || sh.ShotType != -2) && !sh.Stored && (sh.ShotType != -4 || !noWasteShowDecay))
                    sh.Age += 1;

                if (sh.Age > sh.Range && !sh.Flash)
                    _shots.Remove(key);
                else
                    _shots[key] = sh;
            }
        }

        private void ClearRemoved()
        {
            var keys = _shots.Where(s => !s.Value.Exists).Select(s => s.Key).ToList();
            foreach (var key in keys)
                _shots.Remove(key);
        }

        private Shot FieldBorderCollision(Shot shot, bool upDnConnected, bool dxSxConnected, Model.Vector fieldSize)
        {
            if (upDnConnected)
            {
                if (shot.Position.Y > fieldSize.Y)
                    shot.Position.Y -= fieldSize.Y;
                else if (shot.Position.Y < 0)
                    shot.Position.Y += fieldSize.Y;
            }
            else
            {
                if (shot.Position.Y > fieldSize.Y)
                {
                    shot.Position.Y = fieldSize.Y;
                    shot.Velocity.Y = -1 * Math.Abs(shot.Velocity.Y);
                }
                else if (shot.Position.Y < 0)
                {
                    shot.Position.Y = 0;
                    shot.Velocity.Y = Math.Abs(shot.Velocity.Y);
                }
            }
            if (dxSxConnected)
            {
                if (shot.Position.X > fieldSize.X)
                    shot.Position.X -= fieldSize.X;
                else if (shot.Position.Y < 0)
                    shot.Position.X += fieldSize.X;
            }
            else
            {
                if (shot.Position.X > fieldSize.X)
                {
                    shot.Position.X = fieldSize.X;
                    shot.Velocity.X = -1 * Math.Abs(shot.Velocity.X);
                }
                else if (shot.Position.X < 0)
                {
                    shot.Position.X = 0;
                    shot.Velocity.X = Math.Abs(shot.Velocity.X);
                }
            }

            return shot;
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

        private struct Collision
        {
            public Collision(double time, short robot)
            {
                Time = time;
                Robot = robot;
            }

            public short Robot { get; }
            public double Time { get; }
        }
    }
}