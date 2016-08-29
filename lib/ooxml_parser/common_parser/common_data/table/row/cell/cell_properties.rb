require_relative 'merge'
module OoxmlParser
  # Class for parsing 'w:tcPr' element
  class CellProperties < OOXMLDocumentObject
    attr_accessor :fill, :color, :borders, :text_direction, :anchor, :anchor_center, :horizontal_overflow, :merge, :table_cell_width, :borders_properties, :shd, :vertical_align
    # @return [TableMargins] margins
    attr_accessor :table_cell_margin

    def initialize(borders = nil, color = nil)
      @borders = borders
      @color = color
      @fill = fill
      @shd = :none
      @merge = false
    end

    def parse(cell_properties_node)
      @borders_properties = Borders.new
      cell_properties_node.xpath('*').each do |cell_node_child|
        case cell_node_child.name
        when 'vMerge'
          @merge = CellMerge.parse(cell_node_child)
        when 'vAlign'
          @vertical_align = cell_node_child.attribute('val').value.to_sym
        when 'gridSpan'
          @merge = CellMerge.new('horizontal', cell_node_child.attribute('val').value)
        when 'tcW'
          @table_cell_width = cell_node_child.attribute('w').value.to_f / 567.5
        when 'tcMar'
          @table_cell_margin = TableMargins.new(parent: self).parse(cell_node_child)
        when 'textDirection'
          @text_direction = Alignment.parse(cell_node_child.attribute('val'))
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
            @borders_properties.top = BordersProperties.parse(top)
          end
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:right").each do |right|
            @borders_properties.right = BordersProperties.parse(right)
          end
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:left").each do |left|
            @borders_properties.left = BordersProperties.parse(left)
          end
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:bottom").each do |bottom|
            @borders_properties.bottom = BordersProperties.parse(bottom)
          end
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:tl2br").each do |tl2br|
            @borders_properties.top_left_to_bottom_right = BordersProperties.parse(tl2br)
          end
          cell_node_child.xpath("#{OOXMLDocumentObject.namespace_prefix}:tr2bl").each do |tr2bl|
            @borders_properties.top_right_to_bottom_left = BordersProperties.parse(tr2bl)
          end
        end
      end

      cell_properties_node.attributes.each do |key, value|
        case key
        when 'vert'
          @text_direction = value.value.to_sym
        when 'anchor'
          @anchor = Alignment.parse(value)
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
