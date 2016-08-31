module OoxmlParser
  class Indents < OOXMLDocumentObject
    attr_accessor :first_line_indent, :left_indent, :right_indent, :hanging_indent

    def initialize(first_line_indent = OoxmlSize.new(0),
                   left_indent = OoxmlSize.new(0),
                   right_indent = OoxmlSize.new(0),
                   hanging_indent = OoxmlSize.new(0),
                   parent: nil)
      @first_line_indent = first_line_indent
      @left_indent = left_indent
      @right_indent = right_indent
      @hanging_indent = hanging_indent
      @parent = parent
    end

    alias first_line first_line_indent
    alias left left_indent
    alias right right_indent
    alias hanging hanging_indent

    def to_s
      "first line indent: #{@first_line_indent}, left indent: #{@left_indent},
     right indent: #{@right_indent}, hanging indent: #{@hanging_indent}"
    end

    # Parse Indents
    # @param [Nokogiri::XML:Element] node with Indents
    # @return [Indents] value of Indents
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'firstLine'
          @first_line_indent = OoxmlSize.new(value.value.to_f)
        when 'left'
          @left_indent = OoxmlSize.new(value.value.to_f)
        when 'right'
          @right_indent = OoxmlSize.new(value.value.to_f)
        when 'hanging'
          @hanging_indent = OoxmlSize.new(value.value.to_f)
        end
      end
      self
    end
  end
end
