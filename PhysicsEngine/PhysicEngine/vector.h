#pragma once

struct Vector {
	float x;
	float y;

	Vector operator/ (const float scalar) const;
	Vector operator* (const float scalar) const;
	Vector operator+ (const Vector v2) const;
	void operator+=(const Vector v2);
	float MagnitudeSquared() const;
	float Magnitude() const;
	float InverseMagnitude() const;
	Vector Unit() const;
};

