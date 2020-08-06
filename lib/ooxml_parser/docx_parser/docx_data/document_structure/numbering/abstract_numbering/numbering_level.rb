# frozen_string_literal: true

require_relative 'numbering_level/level_justification'
require_relative 'numbering_level/level_text'
require_relative 'numbering_level/numbering_format'
require_relative 'numbering_level/suffix'
module OoxmlParser
  # This element specifies the appearance and behavior of a numbering level
  # within a given abstract numbering
  class NumberingLevel < OOXMLDocumentObject
    # @return [Integer] level id
    attr_accessor :ilvl
    # @return [Start] start data
    attr_accessor :start
    # @return [NumberingFormat] numbering format data
    attr_accessor :numbering_format
    # @return [LevelText] level text
    attr_accessor :text
    # @return [LevelJustification] justification of level
    attr_accessor :justification
    # @return [ParagraphProperties] properties of paragraph
    attr_accessor :paragraph_properties
    # @return [RunProperties] properties of run
    attr_accessor :run_properties
    # @return [Suffix] value of Suffix
    attr_accessor :suffix

    def initialize(parent: nil)
      @suffix = Suffix.new(parent: self)
      super
    end

    # Parse Numbering Level data
    # @param [Nokogiri::XML:Element] node with Numbering Level data
    # @return [NumberingLevel] value of Numbering Level data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'ilvl'
          @ilvl = value.value.to_f
        end
      end

      node.xpath('*').each do |num_level_child|
        case num_level_child.name
        when 'start'
          @start = ValuedChild.new(:integer, parent: self).parse(num_level_child)
        when 'numFmt'
          @numbering_format = NumberingFormat.new(parent: self).parse(num_level_child)
        when 'lvlText'
          @text = LevelText.new(parent: self).parse(num_level_child)
        when 'lvlJc'
          @justification = LevelJustification.new(parent: self).parse(num_level_child)
        when 'pPr'
          @paragraph_properties = ParagraphProperties.new(parent: self).parse(num_level_child)
        when 'rPr'
          @run_properties = RunProperties.new(parent: self).parse(num_level_child)
        when 'suff'
          @suffix = @suffix.parse(num_level_child)
        end
      end

      self
    end
  end
end
