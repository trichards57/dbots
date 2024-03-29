﻿using PhysicsEngine.Model;
using System;
using System.Collections.Generic;

namespace PhysicsEngine.Manager
{
    public interface IBucketManager
    {
        void AddBot(int id, ref Point position);

        Point GetRobotBucket(int id);

        void Initialise(int fieldWidth, int fieldHeight);

        void SetRobotBucket(int id, ref Point bucket);
    }

    public class BucketManager : IBucketManager
    {
        public const int BucketSize = 4000;

        private Bucket[,] _buckets;
        private Dictionary<int, Point> _robotToBucket;

        public void AddBot(int id, ref Point position)
        {
            _buckets[position.X, position.Y].RobotsIds.Add(id);
        }

        public Point GetRobotBucket(int id)
        {
            if (!_robotToBucket.ContainsKey(id))
                _robotToBucket[id] = new Point(-2, -2);

            return _robotToBucket[id];
        }

        public void Initialise(int fieldWidth, int fieldHeight)
        {
            if (_buckets != null)
            {
                for (var i = 0; i < _buckets.GetLength(0); i++)
                {
                    for (var j = 0; j < _buckets.GetLength(1); j++)
                    {
                        _buckets[i, j].AdjacentBuckets.Clear();
                        _buckets[i, j].RobotsIds.Clear();
                    }
                }
            }

            var numXBuckets = (int)Math.Ceiling((float)fieldWidth / BucketSize) + 1;
            var numYBuckets = (int)Math.Ceiling((float)fieldHeight / BucketSize) + 1;

            _buckets = new Bucket[numXBuckets, numYBuckets];

            for (var y = 0; y < numYBuckets; y++)
            {
                for (var x = 0; x < numXBuckets; x++)
                {
                    _buckets[x, y] = new Bucket();

                    if (x > 0)
                        _buckets[x, y].AdjacentBuckets.Add(new Point(x - 1, y));
                    if (x < numXBuckets - 1)
                        _buckets[x, y].AdjacentBuckets.Add(new Point(x + 1, y));
                    if (y > 0)
                        _buckets[x, y].AdjacentBuckets.Add(new Point(x, y - 1));
                    if (y < numYBuckets - 1)
                        _buckets[x, y].AdjacentBuckets.Add(new Point(x, y + 1));
                    if (x > 0 && y > 0)
                        _buckets[x, y].AdjacentBuckets.Add(new Point(x - 1, y - 1));
                    if (x > 0 && y < numYBuckets - 1)
                        _buckets[x, y].AdjacentBuckets.Add(new Point(x - 1, y + 1));
                    if (x < numXBuckets - 1 && y > 0)
                        _buckets[x, y].AdjacentBuckets.Add(new Point(x + 1, y - 1));
                    if (x < numXBuckets - 1 && y < numYBuckets - 1)
                        _buckets[x, y].AdjacentBuckets.Add(new Point(x + 1, y + 1));
                }
            }
        }

        public void RemoveBot(int id, ref Point position)
        {
            _buckets[position.X, position.Y].RobotsIds.Remove(id);
        }

        public void SetRobotBucket(int id, ref Point bucket)
        {
            _robotToBucket[id] = bucket;
        }
    }
}