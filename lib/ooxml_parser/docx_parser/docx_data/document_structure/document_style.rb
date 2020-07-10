# frozen_string_literal: true

require_relative 'document_style/document_style_helper'
module OoxmlParser
  # Class for describing styles containing in +styles.xml+
  class DocumentStyle < OOXMLDocumentObject
    include TableStylePropertiesHelper
    include DocumentStyleHelper

    # @return [True, False] is style default
    attr_reader :default
    # @return [Symbol] Type of style (+:paragraph+ or +:table+)
    attr_accessor :type
    # @return [FixNum] number of style
    attr_accessor :style_id
    # @return [String] name of style
    attr_accessor :name
    # @return [FixNum] id of style on which this style is based
    attr_accessor :based_on
    # @return [FixNum] id of next style
    attr_accessor :next_style
    # @return [DocxParagraphRun] run properties
    attr_accessor :run_properties
    # @return [DocxParagraph] run properties
    attr_accessor :paragraph_properties
    # @return [TableProperties] properties of table
    attr_accessor :table_properties
    # @return [Array, TableStyleProperties] list of table style properties
    attr_accessor :table_style_properties_list
    # @return [TableRowProperties] properties of table row
    attr_accessor :table_row_properties
    # @return [CellProperties] properties of table cell
    attr_accessor :table_cell_properties
    # @return [True, False] Latent Style Primary Style Setting
    # Used to determine if current style is visible in style list in editors
    # According to http://www.wordarticles.com/Articles/WordStyles/LatentStyles.php
    attr_accessor :q_format

    alias visible? q_format

    def initialize(parent: nil)
      @q_format = false
      @table_style_properties_list = []
      @parent = parent
    end

    # @return [String] result of convert of object to string
    def to_s
      "Table style properties list: #{@table_style_properties_list.join(',')}"
    end

    def inspect
      to_s
    end

    # Parse single document style
    # @return [DocumentStyle]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @type = value.value.to_sym
        when 'styleId'
          @style_id = value.value
        when 'default'
          @default = attribute_enabled?(value.value)
        end
      end
      node.xpath('*').each do |subnode|
        case subnode.name
        when 'name'
          @name = subnode.attribute('val').value
        when 'basedOn'
          @based_on = subnode.attribute('val').value
        when 'next'
          @next_style = subnode.attribute('val').value
        when 'rPr'
          @run_properties = DocxParagraphRun.new(parent: self).parse_properties(subnode)
        when 'pPr'
          @paragraph_properties = ParagraphProperties.new(parent: self).parse(subnode)
        when 'tblPr'
          @table_properties = TableProperties.new(parent: self).parse(subnode)
        when 'trPr'
          @table_row_properties = TableRowProperties.new(parent: self).parse(subnode)
        when 'tcPr'
          @table_cell_properties = CellProperties.new(parent: self).parse(subnode)
        when 'tblStylePr'
          @table_style_properties_list << TableStyleProperties.new(parent: self).parse(subnode)
        when 'qFormat'
          @q_format = true
        end
      end
      fill_empty_table_styles
      self
    end
  end
end
