# frozen_string_literal: true

# @author Pavel.Lobashov

require_relative 'spacing/spacing_valued_child'
module OoxmlParser
  # Class to describe spacing
  class Spacing
    # @return [Float] Spacing before paragraph
    attr_accessor :before
    # @return [Float] Spacing after paragraph
    attr_accessor :after
    # @return [Float] Spacing between lines
    attr_accessor :line
    # @return [String] Spacing line rule
    attr_accessor :line_rule
    # @return [LineSpacing] line spacing data
    attr_reader :line_spacing

    def initialize(before = nil, after = 0.35, line = nil, line_rule = nil)
      @before = before
      @after = after
      @line = line
      @line_rule = line_rule
    end

    # Compare this object to other
    # @param other [Object] any other object
    # @return [True, False] result of comparision
    def ==(other)
      self.line_rule = :at_least if line_rule == 'atLeast'
      self.line_rule = :multiple if line_rule == :auto
      other.line_rule = :multiple if other.line_rule == :auto
      self.line_rule = line_rule.to_sym if line_rule.instance_of?(String)

      @before == other.before &&
        @after == other.after &&
        @line == other.line &&
        @line_rule.to_s == other.line_rule.to_s
    end

    # @return [String] result of convert of object to string
    def to_s
      result_string = ''
      variables = instance_variables
      variables.each do |current_variable|
        result_string += "#{current_variable.to_s.sub('@', '')}: #{instance_variable_get(current_variable)}\n"
      end
      result_string
    end

    # Method to copy object
    # @return [Spacing] copied object
    def copy
      Spacing.new(@before, @after, @line, @line_rule)
    end

    # Round value of spacing
    # @param count_of_digits [Integer] how digits to left
    # @return [Spacing] result of round
    def round(count_of_digits = 1)
      before = @before.to_f.round(count_of_digits)
      after = @after.to_f.round(count_of_digits)
      line = @line.to_f.round(count_of_digits)
      Spacing.new(before, after, line, @line_rule)
    end

    # Parse data for Spacing
    # @param [Nokogiri::XML:Element] node with Spacing
    # @return [Nothing]
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'lnSpc'
          @line_spacing = SpacingValuedChild.new(parent: self).parse(node_child)
          self.line = @line_spacing.to_ooxml_size
          self.line_rule = @line_spacing.rule
        when 'spcBef'
          @spacing_before = SpacingValuedChild.new(parent: self).parse(node_child)
          self.before = @spacing_before.to_ooxml_size
        when 'spcAft'
          @spacing_after = SpacingValuedChild.new(parent: self).parse(node_child)
          self.after = @spacing_after.to_ooxml_size
        end
      end
    end

    # Fetch data from `ParagraphSpacing`
    # Which have values with parameters
    # @param valued_spacing [ParagraphSpacing] spacing to get params
    # @return [Spacing]
    def fetch_from_valued_spacing(valued_spacing)
      @before = valued_spacing.before.to_unit(:centimeter).value if valued_spacing.before
      @after = valued_spacing.after.to_unit(:centimeter).value if valued_spacing.after
      @line = valued_spacing.line.to_unit(:centimeter).value if valued_spacing.line
      @line_rule = valued_spacing.line_rule if valued_spacing.line_rule
      self
    end
  end
end
