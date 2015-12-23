# Table Position Data
module OoxmlParser
  class TablePosition
    attr_accessor :left, :right, :top, :bottom, :position_x, :position_y, :horizontal_anchor, :vertical_anchor, :vertical_align_from_anchor,
                  :horizontal_align_from_anchor

    def initialize
      @left = nil
      @right = nil
      @top = nil
      @bottom = nil
      @position_x = nil
      @position_y = nil
    end

    def to_s
      'Table position left: ' + @left.to_s + ', right: ' + @right.to_s + ', top: ' + @top.to_s + ', bottom: ' + @bottom.to_s + 'position x: ' + position_x.to_s + 'position y: ' + position_x.to_s
    end

    def self.parse(node)
      table_p_properties = TablePosition.new
      node.attributes.each do |key, value|
        case key
        when 'leftFromText'
          table_p_properties.left = value.value.to_f / 567.5
        when 'rightFromText'
          table_p_properties.right = value.value.to_f / 567.5
        when 'topFromText'
          table_p_properties.top = value.value.to_f / 567.5
        when 'bottomFromText'
          table_p_properties.bottom = value.value.to_f / 567.5
        when 'tblpX'
          table_p_properties.position_x = (value.value.to_f / 566.9).round(3)
        when 'tblpY'
          table_p_properties.position_y = (value.value.to_f / 566.9).round(3)
        when 'vertAnchor'
          table_p_properties.vertical_anchor = value.value.to_sym
        when 'horzAnchor'
          table_p_properties.horizontal_anchor = value.value.to_sym
        when 'tblpXSpec'
          table_p_properties.horizontal_align_from_anchor = value.value.to_sym
        when 'tblpYSpec'
          table_p_properties.vertical_align_from_anchor = value.value.to_sym
        end
      end
      table_p_properties
    end
  end
end
