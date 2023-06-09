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
    }

}