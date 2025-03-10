# frozen_string_literal: true

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

    # Parse ParagraphSpacing
    # @param [Nokogiri::XML:Node] node with ParagraphSpacing
    # @return [ParagraphSpacing] result of parsing
    def parse(node)
      sorted_attributes(node).each do |key, value|
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

    private

    # This is dirty workaround for situations
    # Then `@line_rule` parsed after `@line` so getting
    # `@line` value is totally screwed up
    # @param [Nokogiri::XML:Node] node with ParagraphSpacing
    # @return [Hash] hash with sorted values
    # TODO: Totally redone parsing of spacing to remove this workaround
    def sorted_attributes(node)
      node.attributes.sort.reverse!.to_h
    end
  end
end
