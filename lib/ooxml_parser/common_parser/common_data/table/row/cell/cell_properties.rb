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
      parse_w_namespace = true
      begin
        cell_properties_node.xpath('w:vMerge') # TODO: Perform not so stupit check
      rescue Nokogiri::XML::XPath::SyntaxError
        parse_w_namespace = false
      end
      if parse_w_namespace
        cell_properties_node.xpath('w:vMerge').each do |v_merge|
          cell_properties.merge = CellMerge.parse(v_merge)
        end
        cell_properties_node.xpath('w:vAlign').each do |valign_node|
          cell_properties.vertical_align = valign_node.attribute('val').value.to_sym
        end
        cell_properties_node.xpath('w:gridSpan').each do |grid_span|
          cell_properties.merge = CellMerge.new('horizontal', grid_span.attribute('val').value)
        end
        cell_properties_node.xpath('w:tcW').each do |table_cell_width|
          cell_properties.table_cell_width = table_cell_width.attribute('w').value.to_f / 567.5
        end
        cell_properties_node.xpath('w:tcMar').each do |cell_margins_node|
          cell_properties.table_cell_margin = TableMargins.parse(cell_margins_node)
        end
        cell_properties_node.xpath('w:textDirection').each do |text_direction|
          cell_properties.text_direction = Alignment.parse(text_direction.attribute('val'))
        end
        cell_properties_node.xpath('w:tcBorders').each do |tc_boders|
          tc_boders.xpath('w:top').each do |top|
            next if top.attribute('val').value == 'nil'
            top_border = BordersProperties.new
            top_border.color = top.attribute('color').value
            if !top.attribute('color').nil? && top.attribute('color').value != 'auto'
              top_border.color = Color.from_int16(top.attribute('color').value)
            end
            top_border.sz = top.attribute('sz').value.to_f / 8.0
            top_border.val = top.attribute('val').value
            top_border.space = top.attribute('space').value
            cell_borders.top = top_border
          end
          tc_boders.xpath('w:right').each do |right|
            next if right.attribute('val').value == 'nil'
            right_border = BordersProperties.new
            right_border.color = right.attribute('color').value
            if !right.attribute('color').nil? && right.attribute('color').value != 'auto'
              right_border.color = Color.from_int16(right.attribute('color').value)
            end
            right_border.sz = right.attribute('sz').value.to_f / 8.0
            right_border.val = right.attribute('val').value
            right_border.space = right.attribute('space').value
            cell_borders.right = right_border
          end
          tc_boders.xpath('w:left').each do |left|
            next if left.attribute('val').value == 'nil'
            left_border = BordersProperties.new
            left_border.color = left.attribute('color').value
            if !left.attribute('color').nil? && left.attribute('color').value != 'auto'
              left_border.color = Color.from_int16(left.attribute('color').value)
            end
            left_border.sz = left.attribute('sz').value.to_f / 8.0
            left_border.val = left.attribute('val').value
            left_border.space = left.attribute('space').value
            cell_borders.left = left_border
          end
          tc_boders.xpath('w:bottom').each do |bottom|
            next if bottom.attribute('val').value == 'nil'
            bottom_border = BordersProperties.new
            bottom_border.color = bottom.attribute('color').value
            if !bottom.attribute('color').nil? && bottom.attribute('color').value != 'auto'
              bottom_border.color = Color.from_int16(bottom.attribute('color').value)
            end
            bottom_border.sz = bottom.attribute('sz').value.to_f / 8.0
            bottom_border.val = bottom.attribute('val').value
            bottom_border.space = bottom.attribute('space').value
            cell_borders.bottom = bottom_border
          end
          tc_boders.xpath('w:tl2br').each do |tl2br|
            top_left_to_bottom_right = BordersProperties.new
            top_left_to_bottom_right.color = tl2br.attribute('color').value
            if !tl2br.attribute('color').nil? && tl2br.attribute('color').value != 'auto'
              top_left_to_bottom_right.color = Color.from_int16(tl2br.attribute('color').value)
            end
            top_left_to_bottom_right.sz = tl2br.attribute('sz').value.to_f / 8.0
            top_left_to_bottom_right.val = tl2br.attribute('val').value
            top_left_to_bottom_right.space = tl2br.attribute('space').value
            cell_borders.top_left_to_bottom_right = top_left_to_bottom_right
          end
          tc_boders.xpath('w:tr2bl').each do |tr2bl|
            top_right_to_bottom_left = BordersProperties.new
            top_right_to_bottom_left.color = tr2bl.attribute('color').value
            if !tr2bl.attribute('color').nil? && tr2bl.attribute('color').value != 'auto'
              top_right_to_bottom_left.color = Color.from_int16(tr2bl.attribute('color').value)
            end
            top_right_to_bottom_left.sz = tr2bl.attribute('sz').value.to_f / 8.0
            top_right_to_bottom_left.val = tr2bl.attribute('val').value
            top_right_to_bottom_left.space = tr2bl.attribute('space').value
            cell_borders.top_right_to_bottom_left = top_right_to_bottom_left
          end
        end
        cell_properties_node.xpath('w:shd').each do |shd|
          cell_properties.shd = shd.attribute('fill').value
          if !shd.attribute('fill').nil? && shd.attribute('fill').value != 'auto'
            cell_properties.shd = Color.from_int16(shd.attribute('fill').value)
          end
        end
      else
        cell_properties_node.xpath('a:fill').each do |fill|
          cell_properties.fill = DocxColorScheme.parse(fill)
        end
      end
      cell_properties.borders_properties = cell_borders
      cell_properties
    end
  end
end
