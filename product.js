const { check, validationResult } = require("express-validator/check");
var multer = require('multer');
const response = require("../Utils/response");
const Product = require('../models/Product')
const productController = require("../controllers/product.ctrl");
const excelController = require("../controllers/excel.ctrl");
const Common = require('./../common/common');
var uploadMultiple = multer({ //multer settings
    storage: Common.uplaodMultipleFiles()
})

var uplaodFileInTemp = multer({ //multer settings
    storage: Common.uplaodFileInTempFolder()
})
module.exports = router => {
    function methodNotAllowedHandler(req, res) {
        res.sendStatus(405);
    }
    /**
     * Add product
     */
    router
        .route('/productAdd')
        .post([
                check("productName").isLength({ min: 2 }).trim().matches(/^[A-Za-z0-9.!@#%^&*()+=:;'""`~<>\s]+$/).withMessage("product name must be valid and at least 2 chars long"),
                check("productLead").optional().matches(/^[a-f\d]{24}$/i).withMessage("Enter Valid user id"),
                check("productTeam").optional().isArray().withMessage("Enter valid team"),
                check("vision").exists().withMessage("vision must be valid"),
                check("strategyType").optional().withMessage("Enter valid strategy type"),
                // check("description").optional().withMessage("Enter Description"),
                // check("tags").optional().withMessage("Enter tags")
            ],
            (req, res) => {
                const errors = validationResult(req);
                if (!errors.isEmpty()) {
                    return res.status(412).json(response.sendError(errors.array()));
                } else {
                    productController.addproduct(req, res)
                }
            })
        .all(methodNotAllowedHandler)

    /**
     * list Product
     */
    router
        .route('/product')
        .get(productController.productList)
        .all(methodNotAllowedHandler)

    /**
     * Delete Product
     */

    router
        .route('/productDelete/:id')
        .delete(productController.productDelete)
        .all(methodNotAllowedHandler)

    /**
     * Add tags
     */
    router
        .route('/productTags')
        .put([
                check("tags").exists().withMessage("tags are required"),
                check("productId").matches(/^[a-f\d]{24}$/i).withMessage("Enter Valid productId")
            ],
            (req, res) => {
                const errors = validationResult(req);
                if (!errors.isEmpty()) {
                    return res.status(412).json(response.sendError(errors.array()));
                } else {
                    productController.addProductTag(req, res)
                }
            })
        .all(methodNotAllowedHandler)

    /**
     * update product details
     */
    router
        .route('/productUpdate')
        .put([
                check("productId").matches(/^[a-f\d]{24}$/i).withMessage("Enter Valid productId"),
                check("productName").optional().isLength({ min: 2 }).matches(/^[A-Za-z0-9.!@#$%^&*()_+='";:<>~`\s]+$/).withMessage("product name must be valid and at least 2 chars long"),
                // check("productLead").optional().matches(/^[a-f\d]{24}$/i).withMessage("Enter Valid user id"),
                // check("productTeam").optional().isArray().withMessage("Enter valid team"),
                check("strategyType").optional().withMessage("Enter valid strategy type"),
            ],
            (req, res) => {

                const errors = validationResult(req);
                if (!errors.isEmpty()) {
                    return res.status(412).json(response.sendError(errors.array()));

                } else if (req.body.strategyType != undefined && req.body.strategyType != "Standalone Product" && req.body.strategyType != "Feature Teams" && req.body.strategyType != "Product Line") {
                    return res.status(400).json(response.sendError("Please enter valid value of strategy type"));
                } else {
                    productController.productDetailUpdate(req, res)
                }
            })
        .all(methodNotAllowedHandler)

    /**
     * update product vision 
     */
    router
        .route('/productVisionUpate')
        .put([
            check("vision").exists().withMessage("vision required"),
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id"),
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.productUpateVision(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /** 
     * add Customer Segments and Competitor Profiles    
     */
    router
        .route('/productMarket')
        .put([
            check('title').exists().matches(/^[^0-9]/)
            .custom((value, { req }) => {
                return new Promise((resolve, reject) => {
                    if (req.body.type == 'customer') {
                        var query = { _id: req.body.productId, 'customerSegments.title': value };
                        var getOneRecord = { 'customerSegments.$': 1 }
                    } else {
                        var query = { _id: req.body.productId, 'competitorProfiles.title': value };
                        var getOneRecord = { 'competitorProfiles.$': 1 }
                    }
                    Product.findOne(query, getOneRecord).exec((err, title) => {
                        if (title !== null) {
                            return reject();
                        } else {
                            return resolve();
                        }
                    });
                });
            }).withMessage('This title is already in use'),
            check("importance").optional().withMessage("Importance required"),
            check("tagline").exists().withMessage("Tagline is required"),
            check("discussion").optional().withMessage("Discussion is required"),
            check("needs").optional().isArray().withMessage("Enter valid needs and desires"),
            check("facts").optional().isArray().withMessage("Enter valid facts"),
            check("needvalue").optional().withMessage("Enter valid value"),
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id"),
            check("type").exists().withMessage("Enter type customer or competitor"),
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.addProductMarket(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /**
     * Edit Product customer and competitor
     */
    router
        .route('/editproductMarket')
        .put([
            check('title').exists().matches(/^[^0-9]/).withMessage('This title is required'),
            check("importance").optional().withMessage("Importance required"),
            check("tagline").exists().withMessage("Tagline is required"),
            check("discussion").optional().withMessage("Discussion is required"),
            check("needs").optional().isArray().withMessage("Enter valid needs and desires"),
            check("facts").optional().isArray().withMessage("Enter valid facts"),
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id"),
            check("id").exists().matches(/^[a-f\d]{24}$/i).withMessage("Please enter valid customer Id"),
            check("type").exists().withMessage("Enter type customer or competitor"),
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.editProductMarket(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /**
     * Add Product Objectives
     */
    router
        .route('/productObjective')
        .put([
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id"),
            check('title').exists().matches(/^[^0-9]/)
            .custom((value, { req }) => {
                return new Promise((resolve, reject) => {
                    Product.findOne({ _id: req.body.productId, 'objectives.title': value }, (err, title) => {
                        if (title !== null) {
                            return reject();
                        } else {
                            return resolve();
                        }
                    });
                });
            }).withMessage('This title is already in use'),
            check("shortName").optional().withMessage("Short name is required"),
            check("importance").optional().withMessage("Importance required"),
            check("discussion").optional().withMessage("Discussion is required"),
            check("kpi").exists().withMessage("Key result is required"),
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.addProductObjectives(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /**
     * Edit Product Objective
     */
    router
        .route('/editProductObjectives')
        .put([
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id"),
            check('title').exists().matches(/^[^0-9]/)
            .custom((value, { req }) => {
                return new Promise((resolve, reject) => {
                    Product.findOne({ _id: req.body.productId, 'objectives.title': value }, { 'objectives.$': 1 }, (err, title) => {
                        if (title !== null && title.objectives[0].title == req.body.title && title.objectives[0]._id != req.body.objectiveId) {
                            return reject();
                        } else {
                            return resolve();
                        }
                    });
                });
            }).withMessage('This title is already in use'),
            check("shortName").optional().withMessage("Short name is required"),
            check("importance").optional().withMessage("Importance required"),
            check("discussion").optional().withMessage("Discussion is required"),
            check("kpi").exists().withMessage("Key result is required"),
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.editProductObjective(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /**
     * Add Product Focus Area
     */
    router
        .route('/productFocusarea')
        .put([
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id"),
            check('title').exists().matches(/^[^0-9]/)
            .custom((value, { req }) => {
                return new Promise((resolve, reject) => {
                    Product.findOne({ _id: req.body.productId, 'focusArea.title': value }, (err, title) => {
                        if (title !== null) {
                            return reject();
                        } else {
                            return resolve();
                        }
                    });
                });
            }).withMessage('This title is already in use'),
            check("importance").optional().withMessage("Importance required"),
            check("shortName").optional().withMessage("Short name is required"),
            check("assignedTo").optional().withMessage("Please assign to contributor"),
            check("description").optional().withMessage("Discussion is required"),
            check("alignedObjective").optional().isArray().withMessage("It must be an array with object id")
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.addProductFocusArea(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /**
     * Edit Product Focus Area
     */
    router
        .route('/editProductFocusarea')
        .put([
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id"),
            check("focusareaId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid focusarea id"),
            check('title').exists().matches(/^[^0-9]/)
            .custom((value, { req }) => {
                return new Promise((resolve, reject) => {
                    Product.findOne({ _id: req.body.productId, 'focusArea.title': value }, { 'focusArea.$': 1 }, (err, title) => {
                        if (title !== null && title.focusArea[0].title == req.body.title && title.focusArea[0]._id != req.body.focusareaId) {
                            return reject();
                        } else {
                            return resolve();
                        }
                    });
                });
            }).withMessage('This title is already in use'),
            check("importance").optional().withMessage("Importance required"),
            check("shortName").optional().withMessage("Short name is required"),
            check("assignedTo").optional().withMessage("Please assign to contributor"),
            check("description").optional().withMessage("Discussion is required"),
            check("alignedObjective").optional().isArray().withMessage("Add aligned Objectives, It must be array")
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.editProductfocusarea(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /**
     * product By id
     */
    router
        .route('/productById')
        .post([
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id")
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.getProductById(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /**
     * Delete Product(customer, competitor, 'objective, facts) 
     */
    router
        .route('/delete')
        .delete([
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id"),
            check("id").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid id")
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.deleteProductsubItem(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /**
     * productName by admin id for add external ideas 
     */

    router
        .route('/productNameList')
        .post([
            check("adminId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid admin id"),
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.ListOfProductName(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /**
     * export single product
     */

    router
        .route('/exportProduct')
        .post([
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id"),
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                productController.exportProductDetail(req, res)
            }
        })
        .all(methodNotAllowedHandler)

    /**
     * import single product(customer, competitor, focusArea, objectives, ideas, projects)
     */

    router
        .route('/importProduct')
        .post(uplaodFileInTemp.single('files'), [
            check("productId").exists().matches(/^[a-f\d]{24}$/i).withMessage("Enter valid product id"),
        ], (req, res) => {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                return res.status(412).json(response.sendError(errors.array()));
            } else {
                excelController.importProductDetail(req, res)
            }
        })
        .all(methodNotAllowedHandler)


    /**
     * Example product data
     * */

    router
        .route('/exampleProduct')
        .get(productController.exampleProductDetail)
        .all(methodNotAllowedHandler)
}