package org.subtlelib.poiooxml.impl.style;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.subtlelib.poiooxml.api.style.AdditiveStyle;
import org.subtlelib.poiooxml.fixtures.AdditiveStyleTestImpl;
import org.subtlelib.poiooxml.fixtures.StyleType;

import com.google.common.collect.ImmutableList;

public class CompositeStyleTest {
    private AdditiveStyle style1type1 = new AdditiveStyleTestImpl("style1", StyleType.type1);

    @Test
    public void testEqualsContract() {
        // given
        CompositeStyle composite1 = new CompositeStyle(ImmutableList.of(style1type1));
        CompositeStyle composite2 = new CompositeStyle(ImmutableList.of(style1type1));

        // verify
        assertEquals(composite1, composite2);
    }
}
