const xlsx = require('xlsx')

// Reading our test file
const file = xlsx.readFile('test.xlsx')

// Sample data set
const customers = [
    {
        name: {
            first: 'niek',
            last: 'last',
        },
        email: 'ab@gmail.com',
        dateOfBirth: "1990-11-01",
        customerId: 1
    },
    {
        name: {
            first: 'salty',
            last: 'last',
        },
        email: 'ab@gmail.com',
        dateOfBirth: "2010-11-01",
        customerId: 1
    },
    {
        name: {
            first: 'jess',
        },
        email: 'ab@gmail.com',
        dateOfBirth: "1990-11-01",
        customerId: 1
    },
    {
        name: {
            first: 'puma',
        },
        email: 'ab@gmail.com',
        dateOfBirth: "1990-11-01",
        customerId: 1
    }
]

const mappedData = customers.map(customer => {
    const { first, last="" } = customer.name;
    const age = 1970 - new Date(new Date(customer.dateOfBirth).getTime() - new Date().getTime()).getFullYear()
    return {
        "First Name" : first,
        "Last Name": last,
        "Email": customer.email,
        "Age":age.toString()
    }
})

const workSheet = xlsx.utils.json_to_sheet(mappedData)

xlsx.utils.book_append_sheet( file, workSheet, "JSON_to_excel" )

// Writing to our file
xlsx.writeFile( file,'test.xlsx')
