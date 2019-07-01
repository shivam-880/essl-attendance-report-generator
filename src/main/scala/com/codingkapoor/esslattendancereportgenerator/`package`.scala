package com.codingkapoor.esslattendancereportgenerator

object `package` {
  def using[A <: {def close() : Unit}, B](resource: A)(f: A => B): B = {
    try {
      f(resource)
    } finally {
      resource.close()
    }
  }
}
