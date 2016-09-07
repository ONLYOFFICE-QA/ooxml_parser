require_relative 'cell_properties/vertical_merge'
require_relative 'merge'
module OoxmlParser
  # Class for parsing 'w:tcPr' element
  class CellProperties < OOXMLDocumentObject
    attr_accessor :fill, :color, :borders, :text_direction, :anchor, :anchor_center, :horizontal_overflow, :table_cell_width, :borders_properties, :shd, :vertical_align
    # @return [GridSpan] data about grid span
    attr_accessor :grid_span
    # @return [TableMargins] margins
    attr_accessor :vertical_merge
    # @return [ParagraphMargins] margins of text in cell
    attr_accessor :margins
    # @return [TableMargins] margins of cell
    attr_accessor :table_cell_margin

    def initialize(borders = nil, color = nil)
      @borders = borders
      @color = color
      @fill = fill
      @shd = :none
    end

    alias table_cell_borders borders_properties

    def parse(cell_properties_node)
      @borders_properties = Borders.new
      @margins = ParagraphMargins.new(parent: self).parse(cell_properties_node)
      cell_properties_node.xpath('*').each do |cell_node_child|
        case cell_node_child.name
        when 'vMerge'
          @vertical_merge = VerticalMerge.new.parse(cell_node_child)
        when 'vAlign'
          @vertical_align = cell_node_child.attribute('val').value.to_sym
        when 'gridSpan'
          @grid_span = GridSpan.new.parse(cell_node_child)
        when 'tcW'
          @table_cell_width = OoxmlSize.new(cell_node_child.attribute('w').value.to_f)
        when 'tcMar'
          @table_cell_margin = TableMargins.new(parent: self).parse(cell_node_child)
        when 'textDirection'
          @text_direction = value_to_symbol(cell_node_child.attribute('val'))
        when 'shd'
          next if cell_node_child.attribute('val').value == 'nil'
          @shd = cell_node_child.attribute('fill').value
          if !cell_node_child.attribute('fill').nil? && cell_node_child.attribute('fill').value != 'auto'
            @shd = Color.from_int16(cell_node_child.attribute('fill').value)
          end
        when 'fill'
          @fill = DocxColorScheme.parse(cell_node_child)
        when 'tcBorders'
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:top").each do |top|
            @borders_properties.top = BordersProperties.new(parent: self).parse(top)
          end
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:right").each do |right|
            @borders_properties.right = BordersProperties.new(parent: self).parse(right)
          end
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:left").each do |left|
            @borders_properties.left = BordersProperties.new(parent: self).parse(left)
          end
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:bottom").each do |bottom|
            @borders_properties.bottom = BordersProperties.new(parent: self).parse(bottom)
          end
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:tl2br").each do |tl2br|
            @borders_properties.top_left_to_bottom_right = BordersProperties.new(parent: self).parse(tl2br)
          end
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:tr2bl").each do |tr2bl|
            @borders_properties.top_right_to_bottom_left = BordersProperties.new(parent: self).parse(tr2bl)
          end
        end
      end

      cell_properties_node.attributes.each do |key, value|
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
