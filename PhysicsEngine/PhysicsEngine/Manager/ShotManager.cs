using PhysicsEngine.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace PhysicsEngine.Manager
{
    [Guid("0E1EEC48-FF16-4E41-B4D8-FEB855A5C019")]
    public interface IShotManager
    {
        [DispId(7)]
        void Clear();

        [DispId(2)]
        int CreateShot();

        [DispId(8)]
        void DeleteShot(int a);

        [DispId(6)]
        int GetMaxShot();

        [DispId(1)]
        Shot GetShot(int i);

        [DispId(4)]
        void LoadShots(string file);

        [DispId(3)]
        void SaveShots(string file);

        [DispId(5)]
        void SetShot(int i, ref Shot value);
    }

    [Guid("E514AE43-3D40-48A3-9C62-D54058826EAA"), ClassInterface(ClassInterfaceType.None)]
    public class ShotManager : IShotManager
    {
        private static readonly Dictionary<int, Shot> _shots = new Dictionary<int, Shot>();

        public ShotManager()
        { }

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

        private void ClearRemoved()
        {
            var keys = _shots.Where(s => !s.Value.Exists).Select(s => s.Key).ToList();
            foreach (var key in keys)
                _shots.Remove(key);
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