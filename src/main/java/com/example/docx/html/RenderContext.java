package com.example.docx.html;

import java.util.Map;

final class RenderContext {
    private final StyleResolver styleResolver;
    private final HyperlinkResolver hyperlinkResolver;
    private final StyleRegistry styleRegistry;
    private final Map<String, String> themeColors;

    RenderContext(StyleResolver styleResolver,
                  HyperlinkResolver hyperlinkResolver,
                  StyleRegistry styleRegistry,
                  Map<String, String> themeColors) {
        this.styleResolver = styleResolver;
        this.hyperlinkResolver = hyperlinkResolver;
        this.styleRegistry = styleRegistry;
        this.themeColors = themeColors;
    }

    StyleResolver styleResolver() {
        return styleResolver;
    }

    HyperlinkResolver hyperlinkResolver() {
        return hyperlinkResolver;
    }

    StyleRegistry styleRegistry() {
        return styleRegistry;
    }

    Map<String, String> themeColors() {
        return themeColors;
    }
}
