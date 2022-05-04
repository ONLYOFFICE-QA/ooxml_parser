# frozen_string_literal: true

module OoxmlParser
  # Class for storing Suffix `w:suff`
  # This element specifies the content which
  # shall be added between a given numbering level's text and the text of
  # every numbered paragraph which references that numbering level.
  # If this element is omitted, then its value shall be assumed to be tab.
  # >ECMA-376, 3rd Edition (June, 2011), Fundamentals and Markup Language Reference 17.9.29.
  class Suffix < OOXMLDocumentObject
    # @return [String] value of suffix
    attr_accessor :value

    def initialize(value = :tab,
                   parent: nil)
      @value = value
      super(parent: parent)
    end

    # Parse Suffix
    # @param [Nokogiri::XML:Node] node with Suffix
    # @return [Suffix] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_sym
        end
      end
      self
    end
  end
end
