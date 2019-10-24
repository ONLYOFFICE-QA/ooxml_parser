# frozen_string_literal: true

# Table Position Data
module OoxmlParser
  class TablePosition < OOXMLDocumentObject
    attr_accessor :left, :right, :top, :bottom, :position_x, :position_y, :horizontal_anchor, :vertical_anchor, :vertical_align_from_anchor,
                  :horizontal_align_from_anchor

    def to_s
      'Table position left: ' + @left.to_s + ', right: ' + @right.to_s + ', top: ' + @top.to_s + ', bottom: ' + @bottom.to_s + 'position x: ' + position_x.to_s + 'position y: ' + position_x.to_s
    end

    # Parse TablePosition object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TablePosition] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'leftFromText'
          @left = OoxmlSize.new(value.value.to_f)
        when 'rightFromText'
          @right = OoxmlSize.new(value.value.to_f)
        when 'topFromText'
          @top = OoxmlSize.new(value.value.to_f)
        when 'bottomFromText'
          @bottom = OoxmlSize.new(value.value.to_f)
        when 'tblpX'
          @position_x = OoxmlSize.new(value.value.to_f)
        when 'tblpY'
          @position_y = OoxmlSize.new(value.value.to_f)
        when 'vertAnchor'
          @vertical_anchor = value.value.to_sym
        when 'horzAnchor'
          @horizontal_anchor = value.value.to_sym
        when 'tblpXSpec'
          @horizontal_align_from_anchor = value.value.to_sym
        when 'tblpYSpec'
          @vertical_align_from_anchor = value.value.to_sym
        end
      end
      self
    end
  end
end
