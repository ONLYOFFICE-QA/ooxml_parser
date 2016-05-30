require_relative 'merge'
module OoxmlParser
  class CellProperties < OOXMLDocumentObject
    attr_accessor :fill, :color, :borders, :margins, :text_direction, :anchor, :anchor_center, :horizontal_overflow, :merge, :table_cell_width, :borders_properties, :shd, :vertical_align

    def initialize(borders = nil, color = nil)
      @borders = borders
      @color = color
      @fill = fill
      @shd = :none
      @merge = false
    end

    alias table_cell_margin margins
    alias table_cell_margin= margins=

    def self.parse(cell_properties_node)
      cell_properties = CellProperties.new(Borders.parse(cell_properties_node), PresentationFill.parse(cell_properties_node))
      cell_properties.margins = ParagraphMargins.parse(cell_properties_node)
      cell_properties_node.attributes.each do |key, value|
        case key
        when 'vert'
          cell_properties.text_direction = value.value.to_sym
        when 'anchor'
          cell_properties.anchor = Alignment.parse(value)
        when 'anchorCtr'
          cell_properties.anchor_center = value.value
        when 'horzOverflow'
          cell_properties.horizontal_overflow = value.value.to_sym
        end
      end
      cell_borders = Borders.new
      cell_properties_node.xpath("#{OOXMLDocumentObject.namespace_prefix}:vMerge").each do |v_merge|
        cell_properties.merge = CellMerge.parse(v_merge)
      end
      cell_properties_node.xpath("#{OOXMLDocumentObject.namespace_prefix}:vAlign").each do |valign_node|
        cell_properties.vertical_align = valign_node.attribute('val').value.to_sym
      end
      cell_properties_node.xpath("#{OOXMLDocumentObject.namespace_prefix}:gridSpan").each do |grid_span|
        cell_properties.merge = CellMerge.new('horizontal', grid_span.attribute('val').value)
      end
      cell_properties_node.xpath("#{OOXMLDocumentObject.namespace_prefix}:tcW").each do |table_cell_width|
        cell_properties.table_cell_width = table_cell_width.attribute('w').value.to_f / 567.5
      end
      cell_properties_node.xpath("#{OOXMLDocumentObject.namespace_prefix}:tcMar").each do |cell_margins_node|
        cell_properties.table_cell_margin = TableMargins.parse(cell_margins_node)
      end
      cell_properties_node.xpath("#{OOXMLDocumentObject.namespace_prefix}:textDirection").each do |text_direction|
        cell_properties.text_direction = Alignment.parse(text_direction.attribute('val'))
      end
      cell_properties_node.xpath("#{OOXMLDocumentObject.namespace_prefix}:tcBorders").each do |tc_boders|
        tc_boders.xpath("#{OOXMLDocumentObject.namespace_prefix}:top").each do |top|
          cell_borders.top = BordersProperties.parse(top)
        end
        tc_boders.xpath("#{OOXMLDocumentObject.namespace_prefix}:right").each do |right|
          cell_borders.right = BordersProperties.parse(right)
        end
        tc_boders.xpath("#{OOXMLDocumentObject.namespace_prefix}:left").each do |left|
          cell_borders.left = BordersProperties.parse(left)
        end
        tc_boders.xpath("#{OOXMLDocumentObject.namespace_prefix}:bottom").each do |bottom|
          cell_borders.bottom = BordersProperties.parse(bottom)
        end
        tc_boders.xpath("#{OOXMLDocumentObject.namespace_prefix}:tl2br").each do |tl2br|
          cell_borders.top_left_to_bottom_right = BordersProperties.parse(tl2br)
        end
        tc_boders.xpath("#{OOXMLDocumentObject.namespace_prefix}:tr2bl").each do |tr2bl|
          cell_borders.top_right_to_bottom_left = BordersProperties.parse(tr2bl)
        end
      end
      cell_properties_node.xpath("#{OOXMLDocumentObject.namespace_prefix}:shd").each do |shd|
        next if shd.attribute('val').value == 'nil'
        cell_properties.shd = shd.attribute('fill').value
        if !shd.attribute('fill').nil? && shd.attribute('fill').value != 'auto'
          cell_properties.shd = Color.from_int16(shd.attribute('fill').value)
        end
      end
      cell_properties_node.xpath("#{OOXMLDocumentObject.namespace_prefix}:fill").each do |fill|
        cell_properties.fill = DocxColorScheme.parse(fill)
      end
      cell_properties.borders_properties = cell_borders
      cell_properties
    end
  end
end
