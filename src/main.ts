const test = () => {
    const user = new User()
    const columns = user.getColumns()
    console.log(columns)
    console.log(User.all())
}
const generateDeliverySchedule = (): void => UseCase.View.GenerateReservationDeliveryPlan.execute()

const confirmDeliverySchedule = (): void => UseCase.View.ConfirmDeliverySchedule.execute()

const onChangeRestaurant = (row: number, restaurant: string): void => UseCase.View.ChangeRestaurant.execute(row, restaurant)
