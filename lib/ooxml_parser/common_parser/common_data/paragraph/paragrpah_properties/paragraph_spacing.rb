module OoxmlParser
  # Class for parsing Paragraph Spacing in paragraph properties `w:spacing`
  class ParagraphSpacing < OOXMLDocumentObject
    # @return [OoxmlSize] value of before spacing
    attr_accessor :before
    # @return [OoxmlSize] value of after spacing
    attr_accessor :after
    # @return [OoxmlSize] value of line spacing
    attr_accessor :line
    # @return [Symbol] value of line rule style
    attr_accessor :line_rule

    # Parse Position
    # @param [Nokogiri::XML:Node] node with Position
    # @return [Position] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'before'
          @before = OoxmlSize.new(value.value.to_f)
        when 'after'
          @after = OoxmlSize.new(value.value.to_f)
        when 'lineRule'
          @line_rule = value.value.sub('atLeast', 'at_least').to_sym
        when 'line'
          @line = if @line_rule == :auto
                    OoxmlSize.new(value.value.to_f, :one_240th_cm)
                  else
                    OoxmlSize.new(value.value.to_f)
                  end
        end
      end
      self
    end
  end
end
