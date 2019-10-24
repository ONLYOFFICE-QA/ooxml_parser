# frozen_string_literal: true

require_relative 'docx_formula/accent'
require_relative 'docx_formula/argument_properties'
require_relative 'docx_formula/bar'
require_relative 'docx_formula/box'
require_relative 'docx_formula/delimeter'
require_relative 'docx_formula/fraction'
require_relative 'docx_formula/function'
require_relative 'docx_formula/group_char'
require_relative 'docx_formula/limit'
require_relative 'docx_formula/math_run'
require_relative 'docx_formula/matrix'
require_relative 'docx_formula/nary'
require_relative 'docx_formula/index'
require_relative 'docx_formula/pre_sub_superscript'
require_relative 'docx_formula/radical'
module OoxmlParser
  # Class for formula data
  class DocxFormula < OOXMLDocumentObject
    attr_accessor :formula_run
    # @return [ArgumentProperties] properties of arguments
    attr_accessor :argument_properties

    def initialize(parent: nil)
      @formula_run = []
      @parent = parent
    end

    # Parse DocxFormula object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxFormula] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'r'
          @formula_run << MathRun.new(parent: self).parse(node_child)
        when 'box', 'borderBox'
          @formula_run << Box.new(parent: self).parse(node_child)
        when 'func'
          @formula_run << Function.new(parent: self).parse(node_child)
        when 'rad'
          @formula_run << Radical.new(parent: self).parse(node_child)
        when 'e', 'eqArr'
          @formula_run << DocxFormula.new(parent: self).parse(node_child)
        when 'nary'
          @formula_run << Nary.new(parent: self).parse(node_child)
        when 'd'
          @formula_run << Delimiter.new(parent: self).parse(node_child)
        when 'sSubSup', 'sSup', 'sSub'
          @formula_run << Index.new(parent: self).parse(node_child)
        when 'f'
          @formula_run << Fraction.new(parent: self).parse(node_child)
        when 'm'
          @formula_run << Matrix.new(parent: self).parse(node_child)
        when 'bar'
          @formula_run << Bar.new(parent: self).parse(node_child)
        when 'acc'
          @formula_run << Accent.new(parent: self).parse(node_child)
        when 'groupChr'
          @formula_run << GroupChar.new(parent: self).parse(node_child)
        when 'argPr'
          @argument_properties = ArgumentProperties.new(parent: self).parse(node_child)
        when 'sPre'
          @formula_run << PreSubSuperscript.new(parent: self).parse(node_child)
        when 'limUpp', 'limLow'
          @formula_run << Limit.new(parent: self).parse(node_child)
        end
      end
      return nil if @formula_run.empty?

      self
    end
  end
end
