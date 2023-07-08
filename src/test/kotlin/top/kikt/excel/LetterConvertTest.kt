package top.kikt.excel

import org.junit.jupiter.api.Test
import kotlin.test.assertEquals

class LetterConvertTest {

    @Test
    fun letterConvert() {
        assertEquals(letterToColumnIndex("A"), 0)
        assertEquals(letterToColumnIndex("B"), 1)
        assertEquals(letterToColumnIndex("I"), 8)
        assertEquals(letterToColumnIndex("Z"), 25)

        assertEquals(letterToColumnIndex("AA"), 26)
        assertEquals(letterToColumnIndex("AB"), 27)
        assertEquals(letterToColumnIndex("AZ"), 51)
        assertEquals(letterToColumnIndex("BA"), 52)
        assertEquals(letterToColumnIndex("BB"), 53)
        assertEquals(letterToColumnIndex("ZZ"), 701)

        assertEquals(letterToColumnIndex("AAA"), 702)
        assertEquals(letterToColumnIndex("AAB"), 703)
        assertEquals(letterToColumnIndex("ABB"), 729)
        assertEquals(letterToColumnIndex("BAA"), 1378)

        assertEquals(letterToColumnIndex("HIT"), 5661)
    }

}