/**
 * StyleBuilder - Fluent API for creating Excel styles
 */

import {
  IStyle,
  IBorder,
  IConditionalFormat,
  IStyleBuilder as IStyleBuilderInterface
} from '../types/style.types';
import { 
  Color, 
  HorizontalAlignment,
  VerticalAlignment,
  BorderStyle, 
  FontStyle 
} from '../types/core.types';

/**
 * StyleBuilder class providing a fluent API for creating Excel styles
 */
export class StyleBuilder implements IStyleBuilderInterface {
  private style: Partial<IStyle> = {};

  constructor() {
    // Configuración por defecto: wrapText true y alineación al centro
    this.style.alignment = {
      horizontal: HorizontalAlignment.CENTER,
      vertical: VerticalAlignment.MIDDLE,
      wrapText: true,
      shrinkToFit: true
    };
  }

  /**
   * Create a new StyleBuilder instance
   */
  static create(): StyleBuilder {
    return new StyleBuilder();
  }

  /**
   * Set font name
   */
  fontName(name: string): StyleBuilder {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.name = name;
    return this;
  }

  /**
   * Set font size
   */
  fontSize(size: number): StyleBuilder {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.size = size;
    return this;
  }

  /**
   * Set font style
   */
  fontStyle(style: FontStyle): StyleBuilder {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.style = style;
    return this;
  }

  /**
   * Set font color
   */
  fontColor(color: Color): StyleBuilder {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.color = color;
    return this;
  }

  /**
   * Make font bold
   */
  fontBold(): StyleBuilder {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.bold = true;
    return this;
  }

  /**
   * Make font italic
   */
  fontItalic(): StyleBuilder {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.italic = true;
    return this;
  }

  /**
   * Make font underlined
   */
  fontUnderline(): StyleBuilder {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.underline = true;
    return this;
  }

  /**
   * Set border on all sides
   */
  border(style: BorderStyle, color?: Color): StyleBuilder {
    if (!this.style.border) {
      this.style.border = {};
    }
    const border: IBorder = { style };
    if (color !== undefined) {
      border.color = color;
    }
    this.style.border.top = border;
    this.style.border.left = border;
    this.style.border.bottom = border;
    this.style.border.right = border;
    return this;
  }

  /**
   * Set top border
   */
  borderTop(style: BorderStyle, color?: Color): StyleBuilder {
    if (!this.style.border) {
      this.style.border = {};
    }
    const border: IBorder = { style };
    if (color !== undefined) {
      border.color = color;
    }
    this.style.border.top = border;
    return this;
  }

  /**
   * Set left border
   */
  borderLeft(style: BorderStyle, color?: Color): StyleBuilder {
    if (!this.style.border) {
      this.style.border = {};
    }
    const border: IBorder = { style };
    if (color !== undefined) {
      border.color = color;
    }
    this.style.border.left = border;
    return this;
  }

  /**
   * Set bottom border
   */
  borderBottom(style: BorderStyle, color?: Color): StyleBuilder {
    if (!this.style.border) {
      this.style.border = {};
    }
    const border: IBorder = { style };
    if (color !== undefined) {
      border.color = color;
    }
    this.style.border.bottom = border;
    return this;
  }

  /**
   * Set right border
   */
  borderRight(style: BorderStyle, color?: Color): StyleBuilder {
    if (!this.style.border) {
      this.style.border = {};
    }
    const border: IBorder = { style };
    if (color !== undefined) {
      border.color = color;
    }
    this.style.border.right = border;
    return this;
  }

  /**
   * Set background color
   */
  backgroundColor(color: Color): StyleBuilder {
    if (!this.style.fill) {
      this.style.fill = { type: 'pattern' };
    }
    this.style.fill.backgroundColor = color;
    this.style.fill.pattern = 'solid';
    return this;
  }

  /**
   * Set horizontal alignment
   */
  horizontalAlign(alignment: HorizontalAlignment): StyleBuilder {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.horizontal = alignment;
    return this;
  }

  /**
   * Set vertical alignment
   */
  verticalAlign(alignment: VerticalAlignment): StyleBuilder {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.vertical = alignment;
    return this;
  }

  /**
   * Center align text
   */
  centerAlign(): StyleBuilder {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.horizontal = HorizontalAlignment.CENTER;
    this.style.alignment.vertical = VerticalAlignment.MIDDLE;
    return this;
  }

  /**
   * Left align text
   */
  leftAlign(): StyleBuilder {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.horizontal = HorizontalAlignment.LEFT;
    return this;
  }

  /**
   * Right align text
   */
  rightAlign(): StyleBuilder {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.horizontal = HorizontalAlignment.RIGHT;
    return this;
  }

  /**
   * Wrap text
   */
  wrapText(): StyleBuilder {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.wrapText = true;
    return this;
  }

  /**
   * Set number format
   */
  numberFormat(format: string): StyleBuilder {
    this.style.numberFormat = format;
    return this;
  }

  /**
   * Set striped rows
   */
  striped(): StyleBuilder {
    this.style.striped = true;
    return this;
  }

  /**
   * Add conditional formatting
   */
  conditionalFormat(format: IConditionalFormat): StyleBuilder {
    if (!this.style.conditionalFormats) {
      this.style.conditionalFormats = [];
    }
    this.style.conditionalFormats.push(format);
    return this;
  }

  /**
   * Build the final style
   */
  build(): IStyle {
    return this.style as IStyle;
  }

  /**
   * Reset the builder
   */
  reset(): StyleBuilder {
    this.style = {};
    // Restaurar configuración por defecto
    this.style.alignment = {
      horizontal: HorizontalAlignment.CENTER,
      vertical: VerticalAlignment.MIDDLE,
      wrapText: true
    };
    return this;
  }

  /**
   * Clone the current style
   */
  clone(): StyleBuilder {
    const cloned = new StyleBuilder();
    cloned.style = JSON.parse(JSON.stringify(this.style));
    return cloned;
  }
} 