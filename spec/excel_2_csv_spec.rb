require 'spec_helper'
require_relative '../services/excel_2_csv'

describe 'Excel2CSV' do
	describe '#run' do
  	before do
  		Excel2CSV.new('data/sample-kmp.xlsx').run
  	end

	  it 'exports data to transaction.csv' do
	  	file = "transactions.csv"
	  	epexct_data_output file
	  end

	  it 'exports data to return.csv' do
	  	file = "return.csv"
	  	epexct_data_output file
	  end

	  it 'exports data to delivered.csv' do
	  	file = "delivered.csv"
	  	epexct_data_output file
	  end

	  it 'exports data to products.csv' do
	  	file = "products.csv"
	  	epexct_data_output file
	  end
	end

	def epexct_data_output file
		test_file = "spec/fixtures/#{file}"
  	output_file =  "data/output/#{file}"

  	test_data = []
  	CSV.foreach(test_file, headers: true) do |row|
		  test_data << row.to_hash
		end

		output_data = []
		CSV.foreach(output_file, headers: true) do |row|
		  output_data << row.to_hash
		end

		test_data.each_with_index do |row, index|
			expect(row).to eq(output_data[index])
		end
	end
end