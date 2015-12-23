# Character Alignment in XLSX
module OoxmlParser
  class XlsxAlignment
    attr_accessor :horizontal, :vertical, :wrap_text, :text_rotation

    def initialize(horizontal = :left, vertical = :bottom, wrap_text = false)
      @horizontal = horizontal
      @vertical = vertical
      @wrap_text = wrap_text
    end

    def self.parse(alignment_node)
      alignment = XlsxAlignment.new
      alignment_node.attributes.each do |key, value|
        case key
        when 'horizontal'
          alignment.horizontal = value.value.to_sym
          alignment.wrap_text = true if alignment.horizontal == :justify
        when 'vertical'
          alignment.vertical = value.value.to_sym
        when 'wrapText'
          alignment.wrap_text = value.value.to_s == '1'
        when 'textRotation'
          alignment.text_rotation = value.value.to_i
        end
      end
      alignment
    end
  end
end
