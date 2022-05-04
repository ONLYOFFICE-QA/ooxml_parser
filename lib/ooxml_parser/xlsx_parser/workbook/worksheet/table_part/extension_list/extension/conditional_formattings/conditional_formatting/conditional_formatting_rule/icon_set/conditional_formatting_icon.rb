# frozen_string_literal: true

module OoxmlParser
  # Class for `cfIcon` data
  class ConditionalFormattingIcon < OOXMLDocumentObject
    # @return [String] Name of icon set
    attr_reader :icon_set
    # @return [Integer] Id of icon in set
    attr_reader :icon_id

    # Parse ConditionalFormattingIcon data
    # @param [Nokogiri::XML:Element] node with ConditionalFormattingIcon data
    # @return [ConditionalFormattingIcon] value of ConditionalFormattingIcon data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'iconSet'
          @icon_set = value.value.to_s
        when 'iconId'
          @icon_id = value.value.to_i
        end
      end
      self
    end
  end
end
