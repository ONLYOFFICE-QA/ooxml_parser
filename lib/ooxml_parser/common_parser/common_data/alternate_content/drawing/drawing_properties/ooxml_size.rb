# frozen_string_literal: true

module OoxmlParser
  # Size of some object
  class OoxmlSize
    # @return [Float] value of size
    attr_accessor :value
    # @return [Symbol] units of measurement
    attr_accessor :unit

    def initialize(value = 0, unit = :dxa)
      @unit = unit
      @value = value
    end

    # Parse OoxmlSize
    # @param [Nokogiri::XML:Node] node with OoxmlSize
    # @return [OoxmlSize] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @unit = value.value.to_sym
        when 'w', 'val'
          @value = value.value.to_f
        end
      end
      self
    end

    def ==(other)
      (to_base_unit.value - other.to_base_unit.value).abs < 10**(OoxmlParser.configuration.accuracy + 2)
    end

    def zero?
      @value.zero?
    end

    # @param unit [Symbol] unit in which output should be
    # @return [String] string representation of size
    def to_s(unit = :centimeter)
      converted = to_unit(unit)
      "#{converted.value} #{converted.unit}"
    end

    # Convert all values to one same base unit
    # @return [OoxmlSize] base unit
    def to_base_unit
      case @unit
      when :centimeter
        OoxmlSize.new(@value * 360_000)
      when :point
        OoxmlSize.new(@value * 12_700)
      when :half_point
        OoxmlSize.new(@value * (12_700 / 2))
      when :one_eighth_point
        OoxmlSize.new(@value * (12_700 / 8))
      when :one_100th_point
        OoxmlSize.new(@value * (12_700 / 100))
      when :one_240th_cm
        OoxmlSize.new(@value * 1500)
      when :dxa, :twip
        OoxmlSize.new(@value * 635, :emu)
      when :inch
        OoxmlSize.new(@value * 914_400, :emu)
      when :spacing_point
        OoxmlSize.new(@value * (12_700 / 100), :emu)
      when :percent
        OoxmlSize.new(@value * 100_000, :one_100000th_percent)
      when :pct
        OoxmlSize.new(@value * 100_000 / 50, :one_100000th_percent)
      when :one_1000th_percent
        OoxmlSize.new(@value * 100, :one_100000th_percent)
      when :one_60000th_degree
        OoxmlSize.new(@value, :one_60000th_degree)
      when :degree
        OoxmlSize.new(@value * 60_000, :one_60000th_degree)
      else
        self
      end
    end

    # @param output_unit [Symbol] output unit of convertion
    # @return [OoxmlSize] converted unit
    def to_unit(output_unit)
      base_unit = to_base_unit
      case output_unit
      when :centimeter
        OoxmlSize.new((base_unit.value / 360_000).round(OoxmlParser.configuration.accuracy), output_unit)
      when :point
        OoxmlSize.new((base_unit.value / 12_700).round(OoxmlParser.configuration.accuracy), output_unit)
      when :half_point
        OoxmlSize.new((base_unit.value / (12_700 * 2)).round(OoxmlParser.configuration.accuracy), output_unit)
      when :one_eighth_point
        OoxmlSize.new((base_unit.value / (12_700 * 8)).round(OoxmlParser.configuration.accuracy), output_unit)
      when :one_100th_point
        OoxmlSize.new((base_unit.value / (12_700 / 100)).round(OoxmlParser.configuration.accuracy), output_unit)
      when :one_240th_cm
        OoxmlSize.new((base_unit.value / 1500).round(OoxmlParser.configuration.accuracy), output_unit)
      when :dxa, :twip
        OoxmlSize.new((base_unit.value / 635).round(OoxmlParser.configuration.accuracy), output_unit)
      when :inch
        OoxmlSize.new((base_unit.value / 914_400).round(OoxmlParser.configuration.accuracy), output_unit)
      when :percent
        OoxmlSize.new((base_unit.value / 50).round(OoxmlParser.configuration.accuracy), output_unit)
      when :spacing_point
        OoxmlSize.new((base_unit.value / (12_700 * 100)).round(OoxmlParser.configuration.accuracy), output_unit)
      else
        base_unit
      end
    end
  end
end
