require_relative 'cell_properties/vertical_merge'
require_relative 'merge'
module OoxmlParser
  # Class for parsing 'w:tcPr' element
  class CellProperties < OOXMLDocumentObject
    attr_accessor :fill, :color, :borders, :text_direction, :anchor, :anchor_center, :horizontal_overflow, :table_cell_width, :borders_properties, :vertical_align
    # @return [GridSpan] data about grid span
    attr_accessor :grid_span
    # @return [TableMargins] margins
    attr_accessor :vertical_merge
    # @return [ParagraphMargins] margins of text in cell
    attr_accessor :margins
    # @return [Shade] shade of cell
    attr_accessor :shade
    # @return [TableMargins] margins of cell
    attr_accessor :table_cell_margin
    # @return [True, False] This element will prevent text from
    # wrapping in the cell under certain conditions. If the cell width is fixed,
    # then noWrap specifies that the cell will not be smaller than that fixed
    # value when other cells in the row are not at their minimum. If the cell
    # width is set to auto or pct, then the content of the cell will not wrap.
    # > ECMA-376, 3rd Edition (June, 2011), Fundamentals and Markup Language Reference 17.4.30.
    attr_accessor :no_wrap

    alias table_cell_borders borders_properties

    def parse(node)
      @borders_properties = Borders.new
      @margins = ParagraphMargins.new(parent: self).parse(node)
      @color = PresentationFill.new(parent: self).parse(node)
      @borders = Borders.new(parent: self).parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'vMerge'
          @vertical_merge = VerticalMerge.new.parse(node_child)
        when 'vAlign'
          @vertical_align = node_child.attribute('val').value.to_sym
        when 'gridSpan'
          @grid_span = GridSpan.new.parse(node_child)
        when 'tcW'
          @table_cell_width = OoxmlSize.new(node_child.attribute('w').value.to_f)
        when 'tcMar'
          @table_cell_margin = TableMargins.new(parent: self).parse(node_child)
        when 'textDirection'
          @text_direction = value_to_symbol(node_child.attribute('val'))
        when 'noWrap'
          @no_wrap = option_enabled?(node_child)
        when 'shd'
          @shade = Shade.new(parent: self).parse(node_child)
        when 'fill'
          @fill = DocxColorScheme.new(parent: self).parse(node_child)
        when 'tcBorders'
          node_child.xpath('*').each do |border_child|
            case border_child.name
            when 'top'
              @borders_properties.top = BordersProperties.new(parent: self).parse(border_child)
            when 'right'
              @borders_properties.right = BordersProperties.new(parent: self).parse(border_child)
            when 'left'
              @borders_properties.left = BordersProperties.new(parent: self).parse(border_child)
            when 'bottom'
              @borders_properties.bottom = BordersProperties.new(parent: self).parse(border_child)
            when 'tl2br'
              @borders_properties.top_left_to_bottom_right = BordersProperties.new(parent: self).parse(border_child)
            when 'tr2bl'
              @borders_properties.top_right_to_bottom_left = BordersProperties.new(parent: self).parse(border_child)
            end
          end
        end
      end

      node.attributes.each do |key, value|
        case key
        when 'vert'
          @text_direction = value.value.to_sym
        when 'anchor'
          @anchor = value_to_symbol(value)
        when 'anchorCtr'
          @anchor_center = value.value
        when 'horzOverflow'
          @horizontal_overflow = value.value.to_sym
        end
      end
      self
    end
  end
end
