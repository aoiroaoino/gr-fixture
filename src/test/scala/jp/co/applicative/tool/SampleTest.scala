package jp.co.applicative.tool

import org.scalatest.FunSpec

class SampleTest extends FunSpec {

  describe("sample test") {
    it("test1") {
      // example failed test
      // assert(FixtureCreator.DOUBLEQUOTE === "")
    }
    it("test2") {
      // example success test
      assert(FixtureCreator.DOUBLEQUOTE === FixtureCreator.DOUBLEQUOTE)
    }
  }

}

