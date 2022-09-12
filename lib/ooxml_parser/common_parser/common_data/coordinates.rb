# frozen_string_literal: true

module OoxmlParser
  # Class for working with coordinates
  class Coordinates
    # @return [Regexp] regexp for row name
    ROW_REGEXP = /[a-z]/i.freeze
    # @return [Regexp] regexp for column name
    COLUMN_REGEXP = /\d/.freeze

    attr_accessor :row, :column, :list

    def initialize(row = nil, column = nil, list = nil)
      @row = row.nil? ? row : row.to_i
      @column = column
      @list = list
    end

    class << self
      # Parse range of coordinates
      # @param arguments_string [String] data to parse
      # @return [Array] result
      def parser_coordinates_range(arguments_string)
        return parse_coordinates_array(arguments_string) if arguments_string.include?(',')
        return warn "Formulas with # is unsupported: #{arguments_string}" if arguments_string.include?('#')
        return warn 'Formulas consists from `!` only' if arguments_string == '!'

        sheet_name = 'Sheet1'

        if arguments_string.include?('!')
          sheet_name = arguments_string.match(/.*!/).to_s
          arguments_string = arguments_string.sub(sheet_name, '')
        end

        range = arguments_string.split(':')

        difference = []
        symbols_from = range.first.scan(ROW_REGEXP).join
        symbols_to = range.last.scan(ROW_REGEXP).join
        digits_from = range.first.scan(COLUMN_REGEXP).join
        digits_to = range.last.scan(COLUMN_REGEXP).join

        difference[0] = [symbols_from, symbols_to] unless symbols_from == symbols_to
        difference[1] = [digits_from, digits_to] unless digits_from == digits_to
        arguments_array = []
        case difference.length
        when 0
          arguments_array << Coordinates.new(digits_from, symbols_from)
        when 1
          (difference.first.first..difference.first.last).each do |symbol|
            arguments_array << Coordinates.new(digits_from, symbol, sheet_name)
          end
        when 2
          case difference.first
          when nil
            (difference.last.first..difference.last.last).each do |digit|
              arguments_array << Coordinates.new(digit, symbols_from, sheet_name)
            end
          else
            (difference.first.first..difference.first.last).each do |symbol|
              (difference.last.first..difference.last.last).each do |digit|
                arguments_array << Coordinates.new(digit, symbol, sheet_name)
              end
            end
          end
        end
        arguments_array
      end

      # Parse array of coordinates
      # @param arguments_string [String] string
      # @return [Array] result
      def parse_coordinates_array(arguments_string)
        result = []
        coord_array = arguments_string.split(',')
        coord_array.each do |current_coord|
          result << parser_coordinates_range(current_coord)
        end
        result
      end

      # This method check is argument contains coordinate
      # @param [String] string
      def coordinates?(string)
        !/^([A-Z]+)(\d+)$/.match(string).nil?
      end

      # @param [String or Fixnum] number to convert
      # @return [String] column's name
      # This method takes string like '12' or '45' etc.
      # and converts into spreadsheet column's name
      #    StringHelper.get_column_name('12')  #=> "L"
      #    StringHelper.get_column_name('45')  #=> "AS"
      #    StringHelper.get_column_name('287')  #=> "KA"
      def get_column_name(number)
        (number.to_i.positive? ? ('A'..'Z').to_a[(number.to_i - 1) % 26] + get_column_name((number.to_i - 1) / 26).reverse : '').reverse
      end
    end

    # Parse string to coordinates
    # @param [String] string to parse
    # @return [Coordinates] result
    def parse_string(string)
      string = extract_list_from_string(string) if list_name_in_string?(string)
      if Coordinates.coordinates?(string)
        @row = string.scan(/\d/).join.to_i
        @column = string.scan(/[A-Z]/).join
      end
      self
    end

    # @return [Integer] number of column
    # This method takes @column string
    # and converts into integer
    def column_number
      @column.reverse.each_char.reduce(0) do |result, char|
        result + ((char.downcase.ord - 96) * (26**@column.reverse.index(char)))
      end
    end

    # Compares rows of two cells
    # @param [Coordinates] other_cell other cell coordinates
    # @return [true, false] true, if row greater, than other row
    def row_greater_that_other?(other_cell)
      @row > other_cell.row
    end

    # Compares columns of two cells
    # @param [Coordinates] other_cell other cell coordinates
    # @return [true, false] true, if column greater, than other row
    def column_greater_that_other?(other_cell)
      column_number > other_cell.column_number
    end

    # @return [String] result of convert of object to string
    def to_s
      "#{@column}#{@row} #{@list ? "list: #{@list}" : ''}"
    end

    def nil?
      @column.nil? && @list.nil? && @row.nil?
    end

    # Compare this object to other
    # @param other [Object] any other object
    # @return [True, False] result of comparision
    def ==(other)
      other.is_a?(Coordinates) ? (@row == other.row && @column == other.column) : false
    end

    private

    # Check if there is list name in string
    # @param [String] string to check
    # @return [Boolean] result
    def list_name_in_string?(string)
      string.include?('!')
    end

    # Extract list name and leave coordinates
    # @param [String] string to parse
    # @return [String] coordinates string without list
    def extract_list_from_string(string)
      split = string.split('!')
      @list = split[0...-1].join('!')
      split[-1]
    end
  end
end
