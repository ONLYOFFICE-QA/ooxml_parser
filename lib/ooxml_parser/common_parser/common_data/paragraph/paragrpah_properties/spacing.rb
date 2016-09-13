# endcoding: utf-8
# @author Pavel.Lobashov

# Class to describe spacing
module OoxmlParser
  class Spacing
    # @return [Float] Spacing before paragraph
    attr_accessor :before
    # @return [Float] Spacing after paragraph
    attr_accessor :after
    # @return [Float] Spacing between lines
    attr_accessor :line
    # @return [String] Spacing line rule
    attr_accessor :line_rule

    def initialize(before = nil, after = 0.35, line = nil, line_rule = nil)
      @before = before
      @after = after
      @line = line
      @line_rule = line_rule
    end

    def ==(other)
      self.line_rule = :at_least if line_rule == 'atLeast'
      self.line_rule = :multiple if line_rule == :auto
      other.line_rule = :multiple if other.line_rule == :auto
      if self.class == NilClass || other.class == NilClass
        return true if self.class == NilClass && other.class == NilClass
        return false
      else
        self.line_rule = line_rule.to_sym if line_rule.instance_of?(String)

        if @before == other.before &&
           @after == other.after &&
           @line == other.line &&
           @line_rule.to_s == other.line_rule.to_s
          return true
        else
          return false
        end
      end
    end

    def to_s
      result_string = ''
      variables = instance_variables
      variables.each do |current_variable|
        result_string += current_variable.to_s.sub('@', '') + ': ' + instance_variable_get(current_variable).to_s + "\n"
      end
      result_string
    end

    def copy
      Spacing.new(@before, @after, @line, @line_rule)
    end

    def self.default_spacing_canvas(line_spacing = 1.15)
      Spacing.new(0.0, 0.35277777777777775, line_spacing)
    end

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
      node.xpath('*').each do |spacing_node_child|
        case spacing_node_child.name
        when 'lnSpc'
          self.line = Spacing.parse_spacing_value(spacing_node_child)
          self.line_rule = Spacing.parse_spacing_rule(spacing_node_child)
        when 'spcBef'
          self.before = Spacing.parse_spacing_value(spacing_node_child)
        when 'spcAft'
          self.after = Spacing.parse_spacing_value(spacing_node_child)
        end
      end
    end

    # Parse values of spacing number
    # @param [Nokogiri::XML:Element] node with Spacing
    # @return [Float] value of spacing number
    def self.parse_spacing_value(node)
      node.xpath('*').each do |spacing_node_child|
        case spacing_node_child.name
        when 'spcPct'
          return OoxmlSize.new(spacing_node_child.attribute('val').value.to_f, :one_1000th_percent)
        when 'spcPts'
          return OoxmlSize.new(spacing_node_child.attribute('val').value.to_f, :spacing_point)
        end
      end
    end

    # Parse spacing rule
    # @param [Nokogiri::XML:Element] node with Spacing
    # @return [Symbol] type of spacing rule
    def self.parse_spacing_rule(node)
      node.xpath('*').each do |spacing_node_child|
        case spacing_node_child.name
        when 'spcPct'
          return :multiple
        when 'spcPts'
          return :exact
        else
          return :at_least
        end
      end
    end
  end
end
